import { ChangeEvent, useEffect, useMemo, useRef, useState } from 'react'
import {
  CalendarDays,
  CheckCircle2,
  ChevronLeft,
  CircleAlert,
  FileSpreadsheet,
  History,
  Plus,
  Printer,
  Search,
  ShieldCheck,
  Upload,
  Users,
  WandSparkles,
  UserRoundCheck,
  Trash2,
  LogOut,
  X,
} from 'lucide-react'
import * as XLSX from 'xlsx'
import { Nav } from './components/Nav'
import { supabase } from './lib/supabase'
import { startingActivities, activityNameFromList } from './data/activities'
import { startingStaff } from './data/staff'
import type {
  Page,
  ArrivalAssignment,
  ProgrammeImport,
  ProgrammeRow,
  StaffMember,
  StaffRole,
  StaffingAssignment,
  Activity,
} from './types'

const PROGRAMME_KEY = 'acm-programme-current'
const HISTORY_KEY = 'acm-programme-history'
const STAFF_KEY = 'acm-staff'
const ASSIGNMENT_KEY = 'acm-assignments'
const SICKNESS_KEY = 'acm-sickness-by-day'
const WORKING_KEY = 'acm-working-by-day'
const ACTIVITIES_KEY = 'acm-activities'
const ARRIVAL_ASSIGNMENTS_KEY = 'acm-arrival-assignments'

function readJson<T>(key: string, fallback: T): T {
  try {
    const value = localStorage.getItem(key)
    return value ? JSON.parse(value) : fallback
  } catch {
    return fallback
  }
}

function normaliseText(value: unknown) {
  return String(value ?? '').trim()
}

function isSessionValue(value: unknown) {
  const text = normaliseText(value)
  return /^[1-9]\d*$/.test(text)
}

function parseProgrammeWorkbook(
  workbook: XLSX.WorkBook,
  fileName: string,
): ProgrammeImport {
  const candidateSheets = workbook.SheetNames.map((sheetName) => {
    const rows = XLSX.utils.sheet_to_json<unknown[]>(
      workbook.Sheets[sheetName],
      { header: 1, raw: false, defval: '' },
    )
    const dayRow = rows.findIndex(
      (row) =>
        normaliseText(row?.[0]).toUpperCase() === 'DAY' &&
        normaliseText(row?.[1]).toUpperCase() === 'SES',
    )
    return { sheetName, rows, dayRow }
  }).filter((candidate) => candidate.dayRow >= 0)

  if (!candidateSheets.length) {
    throw new Error(
      'I could not find the DAY / SES programme table in this workbook.',
    )
  }

  const candidate = candidateSheets.sort((a, b) => {
    const aScore = a.rows.flat().filter((value) => normaliseText(value)).length
    const bScore = b.rows.flat().filter((value) => normaliseText(value)).length
    return bScore - aScore
  })[0]

  const { sheetName, rows, dayRow } = candidate
  const groupHeaderRow = rows[dayRow + 1] ?? []

  const groupColumns: { column: number; group: number }[] = []
  for (let column = 2; column < groupHeaderRow.length; column += 1) {
    const parsed = Number(normaliseText(groupHeaderRow[column]))
    if (Number.isInteger(parsed) && parsed >= 1 && parsed <= 30) {
      groupColumns.push({ column, group: parsed })
    }
  }

  if (!groupColumns.length) {
    throw new Error('No numbered group columns were found in the programme.')
  }

  let currentDay = ''
  let pendingSchoolLabel = ''
  const programmeRows: ProgrammeRow[] = []

  for (let rowIndex = dayRow + 2; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] ?? []
    const dayValue = normaliseText(row[0]).toUpperCase()
    if (dayValue && /^[A-Z]{3,9}$/.test(dayValue)) currentDay = dayValue

    const sessionValue = normaliseText(row[1])

    if (!isSessionValue(sessionValue)) {
      const possibleSchool = normaliseText(row[groupColumns[0].column])
      const hasSeveralActivities = groupColumns.filter(({ column }) =>
        /^[A-Z][A-Z0-9 ]{0,10}$/.test(normaliseText(row[column])),
      ).length

      if (
        possibleSchool &&
        possibleSchool.length > 3 &&
        hasSeveralActivities < 3 &&
        !/^(YR|YEAR)\b/i.test(possibleSchool)
      ) {
        pendingSchoolLabel = possibleSchool
      }
      continue
    }

    const cells = groupColumns.map(({ column, group }) => ({
      group,
      activityCode: normaliseText(row[column]).toUpperCase(),
    }))

    const populatedCells = cells.filter((cell) => cell.activityCode)
    if (!populatedCells.length) continue

    programmeRows.push({
      id: `${currentDay}-${sessionValue}-${rowIndex}`,
      day: currentDay || 'DAY',
      session: sessionValue,
      schoolLabel: pendingSchoolLabel || undefined,
      cells,
    })
    pendingSchoolLabel = ''
  }

  if (!programmeRows.length) {
    throw new Error('No programme sessions could be read from the workbook.')
  }

  const title =
    rows
      .slice(0, Math.max(dayRow, 1))
      .flat()
      .map(normaliseText)
      .find((value) => value.length > 4) ?? fileName

  return {
    title,
    sheetName,
    groupNumbers: groupColumns.map((item) => item.group),
    rows: programmeRows,
    importedAt: new Date().toISOString(),
    sourceFileName: fileName,
  }
}

function cellKey(rowId: string, group: number) {
  return `${rowId}::${group}`
}

function ManagerApp({
  accountEmail,
  onSignOut,
}: {
  accountEmail: string
  onSignOut: () => void
}) {
  const [page, setPage] = useState<Page>('dashboard')
  const [programme, setProgramme] = useState<ProgrammeImport | null>(() =>
    readJson(PROGRAMME_KEY, null),
  )
  const [history, setHistory] = useState<ProgrammeImport[]>(() =>
    readJson(HISTORY_KEY, []),
  )
  const [staff, setStaff] = useState<StaffMember[]>(() => {
    const saved = readJson<StaffMember[]>(STAFF_KEY, [])
    if (!saved.length) return startingStaff

    const savedByName = new Map(
      saved.map((member) => [member.name.trim().toLowerCase(), member]),
    )

    const merged = startingStaff.map((seedMember) => {
      const savedMember = savedByName.get(
        seedMember.name.trim().toLowerCase(),
      )
      return savedMember
        ? {
            ...seedMember,
            ...savedMember,
            qualifications:
              savedMember.qualifications?.length
                ? savedMember.qualifications
                : seedMember.qualifications,
            signOffs:
              savedMember.signOffs &&
              Object.keys(savedMember.signOffs).length
                ? savedMember.signOffs
                : seedMember.signOffs,
          }
        : seedMember
    })

    const importedNames = new Set(
      startingStaff.map((member) => member.name.trim().toLowerCase()),
    )
    const customStaff = saved.filter(
      (member) => !importedNames.has(member.name.trim().toLowerCase()),
    )

    const next = [...merged, ...customStaff]
    localStorage.setItem(STAFF_KEY, JSON.stringify(next))
    return next
  })
  const [activities, setActivities] = useState<Activity[]>(() =>
    readJson(ACTIVITIES_KEY, startingActivities),
  )
  const [assignments, setAssignments] = useState<StaffingAssignment>(() =>
    readJson(ASSIGNMENT_KEY, {}),
  )
  const [sicknessByDay, setSicknessByDay] = useState<Record<string, string[]>>(() =>
    readJson(SICKNESS_KEY, {}),
  )
  const [workingByDay, setWorkingByDay] = useState<Record<string, string[]>>(() =>
    readJson(WORKING_KEY, {}),
  )
  const [arrivalAssignments, setArrivalAssignments] = useState<
    Record<string, ArrivalAssignment>
  >(() => readJson(ARRIVAL_ASSIGNMENTS_KEY, {}))
  const [selectedCell, setSelectedCell] = useState<{
    row: ProgrammeRow
    group: number
  } | null>(null)
  const [selectedStaffingCell, setSelectedStaffingCell] = useState<{
    row: ProgrammeRow
    group: number
  } | null>(null)
  const [query, setQuery] = useState('')
  const programmeDays = useMemo(
    () => Array.from(new Set(programme?.rows.map((row) => row.day) ?? [])),
    [programme],
  )
  const [selectedStaffingDay, setSelectedStaffingDay] = useState('')
  const activeStaffingDay =
    selectedStaffingDay && programmeDays.includes(selectedStaffingDay)
      ? selectedStaffingDay
      : programmeDays[0] ?? ''
  const [importMessage, setImportMessage] = useState('')
  const [cloudSyncStatus, setCloudSyncStatus] = useState<
    'idle' | 'syncing' | 'synced' | 'error'
  >('idle')
  const syncTimeoutRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const [showSickPanel, setShowSickPanel] = useState(false)
  const [showWorkingPanel, setShowWorkingPanel] = useState(false)
  const [showActivityManager, setShowActivityManager] = useState(false)
  const [newActivityCode, setNewActivityCode] = useState('')
  const [newActivityName, setNewActivityName] = useState('')
  const [arrivalStaffGroup, setArrivalStaffGroup] = useState<
    'all' | StaffRole
  >('all')
  const [showAddStaff, setShowAddStaff] = useState(false)
  const [newStaffName, setNewStaffName] = useState('')
  const [newStaffRole, setNewStaffRole] = useState<StaffRole>('staff')
  const [newStaffQualifications, setNewStaffQualifications] = useState<string[]>([])
  const fileInputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    if (!programme) return
    ensureWorkingStaffForDays(
      Array.from(new Set(programme.rows.map((row) => row.day))),
    )
  }, [programme, staff])

  function sessionDemandForDay(day: string) {
    if (!programme) return []

    const sessionMap = new Map<
      string,
      { session: string; activityStaff: number; arrivalStaff: number }
    >()

    for (const row of programme.rows.filter((item) => item.day === day)) {
      const activeCells = row.cells.filter(
        (cell) => cell.activityCode && cell.activityCode !== 'Z',
      )

      const current = sessionMap.get(row.session) ?? {
        session: row.session,
        activityStaff: 0,
        arrivalStaff: 0,
      }

      current.activityStaff += activeCells.length

      if (row.session === '3' && row.schoolLabel) {
        const groupCount = activeCells.length
        // One team leader plus one instructor per group.
        current.arrivalStaff += groupCount > 0 ? groupCount + 1 : 0
      }

      sessionMap.set(row.session, current)
    }

    return Array.from(sessionMap.values())
      .map((item) => ({
        ...item,
        total: Math.max(item.activityStaff, item.arrivalStaff),
      }))
      .sort((a, b) => Number(a.session) - Number(b.session))
  }

  function busiestSessionForDay(day: string) {
    return sessionDemandForDay(day).reduce<
      { session: string; total: number } | null
    >((busiest, item) => {
      if (!busiest || item.total > busiest.total) {
        return { session: item.session, total: item.total }
      }
      return busiest
    }, null)
  }

  function busiestSessionAcrossProgramme() {
    if (!programme) return null

    return programmeDays.reduce<
      { day: string; session: string; total: number } | null
    >((busiest, day) => {
      const current = busiestSessionForDay(day)
      if (!current) return busiest
      if (!busiest || current.total > busiest.total) {
        return { day, session: current.session, total: current.total }
      }
      return busiest
    }, null)
  }

  const populatedCells = useMemo(
    () =>
      programme?.rows.flatMap((row) =>
        row.cells
          .filter((cell) => cell.activityCode && cell.activityCode !== 'Z')
          .map((cell) => ({ row, cell })),
      ) ?? [],
    [programme],
  )

  const assignedCount = populatedCells.filter(({ row, cell }) =>
    assignments[cellKey(row.id, cell.group)],
  ).length


  function resolvedRole(member: StaffMember): StaffRole {
    if (member.role) return member.role
    return member.teamLeader ? 'teamLeader' : 'staff'
  }

  function roleLabel(role: StaffRole) {
    if (role === 'teamLeader') return 'Team leader'
    if (role === 'activityManager') return 'Activities manager'
    if (role === 'centreManager') return 'Centre manager'
    return 'Staff'
  }

  function rolePriority(role: StaffRole) {
    if (role === 'staff') return 0
    if (role === 'teamLeader') return 1
    if (role === 'activityManager') return 2
    return 3
  }

  function activityName(code: string) {
    return activityNameFromList(activities, code)
  }

  function ensureWorkingStaffForDays(days: string[]) {
    setWorkingByDay((current) => {
      let changed = false
      const next = { ...current }
      const allStaffIds = staff.map((member) => member.id)

      for (const day of days) {
        if (!next[day]) {
          next[day] = allStaffIds
          changed = true
          continue
        }

        // Any newly imported or newly added staff default to working.
        const merged = Array.from(new Set([...next[day], ...allStaffIds]))
        if (merged.length !== next[day].length) {
          next[day] = merged
          changed = true
        }
      }

      if (changed) {
        localStorage.setItem(WORKING_KEY, JSON.stringify(next))
      }

      return changed ? next : current
    })
  }

  function saveProgramme(next: ProgrammeImport, previous?: ProgrammeImport) {
    const nextHistory = previous
      ? [previous, ...history].slice(0, 12)
      : history
    setProgramme(next)
    setHistory(nextHistory)
    localStorage.setItem(PROGRAMME_KEY, JSON.stringify(next))
    localStorage.setItem(HISTORY_KEY, JSON.stringify(nextHistory))
  }

  async function importExcel(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return

    try {
      setImportMessage('Reading programme…')
      const data = await file.arrayBuffer()
      const workbook = XLSX.read(data, { type: 'array' })
      const imported = parseProgrammeWorkbook(workbook, file.name)
      saveProgramme(imported, programme ?? undefined)
      ensureWorkingStaffForDays(
        Array.from(new Set(imported.rows.map((row) => row.day))),
      )
      setImportMessage(
        `Imported ${imported.rows.length} sessions from ${imported.sheetName}.`,
      )
      setPage('programme')
    } catch (error) {
      setImportMessage(
        error instanceof Error ? error.message : 'The programme could not be imported.',
      )
    }
  }

  function restoreVersion(version: ProgrammeImport) {
    saveProgramme(
      {
        ...version,
        importedAt: new Date().toISOString(),
      },
      programme ?? undefined,
    )
    setImportMessage(`Restored ${version.sourceFileName}.`)
  }

  function updateActivity(rowId: string, group: number, activityCode: string) {
    if (!programme) return
    const next = {
      ...programme,
      rows: programme.rows.map((row) =>
        row.id === rowId
          ? {
              ...row,
              cells: row.cells.map((cell) =>
                cell.group === group
                  ? { ...cell, activityCode }
                  : cell,
              ),
            }
          : row,
      ),
    }
    setProgramme(next)
    localStorage.setItem(PROGRAMME_KEY, JSON.stringify(next))
    setSelectedCell(null)

    const key = cellKey(rowId, group)
    if (assignments[key]) {
      const nextAssignments = { ...assignments }
      delete nextAssignments[key]
      setAssignments(nextAssignments)
      localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(nextAssignments))
    }
  }

  function toggleQualification(staffId: string, code: string) {
    const next = staff.map((member) => {
      if (member.id !== staffId) return member
      const has = member.qualifications.includes(code)
      const nextSignOffs = { ...(member.signOffs ?? {}) }
      if (has) delete nextSignOffs[code]
      else nextSignOffs[code] = 'X'
      return {
        ...member,
        qualifications: has
          ? member.qualifications.filter((item) => item !== code)
          : [...member.qualifications, code],
        signOffs: nextSignOffs,
      }
    })
    setStaff(next)
    localStorage.setItem(STAFF_KEY, JSON.stringify(next))
  }

  function selectAllQualifications(staffId: string) {
    const allCodes = activities.map((activity) => activity.code)
    const next = staff.map((member) =>
      member.id === staffId
        ? {
            ...member,
            qualifications: allCodes,
            signOffs: Object.fromEntries(
              allCodes.map((code) => [
                code,
                member.signOffs?.[code] ?? 'X',
              ]),
            ),
          }
        : member,
    )
    setStaff(next)
    localStorage.setItem(STAFF_KEY, JSON.stringify(next))
  }

  function clearAllQualifications(staffId: string) {
    const next = staff.map((member) =>
      member.id === staffId
        ? { ...member, qualifications: [], signOffs: {} }
        : member,
    )
    setStaff(next)
    localStorage.setItem(STAFF_KEY, JSON.stringify(next))
  }

  function setStaffRole(staffId: string, role: StaffRole) {
    const next = staff.map((member) =>
      member.id === staffId
        ? { ...member, role, teamLeader: role === 'teamLeader' }
        : member,
    )
    setStaff(next)
    localStorage.setItem(STAFF_KEY, JSON.stringify(next))
  }

  function updateStaffEmail(staffId: string, email: string) {
    const next = staff.map((member) =>
      member.id === staffId
        ? { ...member, email: email.trim().toLowerCase() }
        : member,
    )
    setStaff(next)
    localStorage.setItem(STAFF_KEY, JSON.stringify(next))
  }

  function addStaffMember() {
    const trimmedName = newStaffName.trim()
    if (!trimmedName) {
      setImportMessage('Enter the new staff member’s name.')
      return
    }

    const nextMember: StaffMember = {
      id: crypto.randomUUID(),
      name: trimmedName,
      role: newStaffRole,
      teamLeader: newStaffRole === 'teamLeader',
      qualifications: newStaffQualifications,
      signOffs: Object.fromEntries(
        newStaffQualifications.map((code) => [code, 'X']),
      ),
    }

    const next = [...staff, nextMember]
    setStaff(next)
    localStorage.setItem(STAFF_KEY, JSON.stringify(next))
    setNewStaffName('')
    setNewStaffRole('staff')
    setNewStaffQualifications([])
    setShowAddStaff(false)
    setImportMessage(`${trimmedName} was added to the staff platform.`)
  }

  function deleteStaffMember(staffId: string) {
    const member = staff.find((item) => item.id === staffId)
    if (!member) return
    if (!window.confirm(`Remove ${member.name} from the platform?`)) return

    const nextStaff = staff.filter((item) => item.id !== staffId)
    const nextAssignments = Object.fromEntries(
      Object.entries(assignments).filter(([, assignedId]) => assignedId !== staffId),
    )
    const nextSickness = Object.fromEntries(
      Object.entries(sicknessByDay).map(([day, ids]) => [
        day,
        ids.filter((id) => id !== staffId),
      ]),
    )

    setStaff(nextStaff)
    setAssignments(nextAssignments)
    setSicknessByDay(nextSickness)
    localStorage.setItem(STAFF_KEY, JSON.stringify(nextStaff))
    localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(nextAssignments))
    localStorage.setItem(SICKNESS_KEY, JSON.stringify(nextSickness))
  }

  function toggleNewStaffQualification(code: string) {
    setNewStaffQualifications((current) =>
      current.includes(code)
        ? current.filter((item) => item !== code)
        : [...current, code],
    )
  }

  function clearDayStaffing(day: string) {
    if (!programme || !day) return

    if (
      !window.confirm(
        `Remove every activity and school-arrival assignment for ${day}?`,
      )
    ) {
      return
    }

    const dayRowIds = new Set(
      programme.rows
        .filter((row) => row.day === day)
        .map((row) => row.id),
    )

    const nextAssignments = Object.fromEntries(
      Object.entries(assignments).filter(([key]) => {
        const [rowId] = key.split('::')
        return !dayRowIds.has(rowId)
      }),
    )

    const nextArrivalAssignments = Object.fromEntries(
      Object.entries(arrivalAssignments).filter(([key]) => {
        const [assignmentDay] = key.split('::')
        return assignmentDay !== day
      }),
    )

    setAssignments(nextAssignments)
    setArrivalAssignments(nextArrivalAssignments)
    localStorage.setItem(
      ASSIGNMENT_KEY,
      JSON.stringify(nextAssignments),
    )
    localStorage.setItem(
      ARRIVAL_ASSIGNMENTS_KEY,
      JSON.stringify(nextArrivalAssignments),
    )
    setImportMessage(`All staffing assignments were cleared for ${day}.`)
  }

  function autoFillStaffing(day: string) {
    if (!programme || !day) return

    const sickIds = new Set(sicknessByDay[day] ?? [])
    const workingIds = new Set(
      workingByDay[day] ?? staff.map((member) => member.id),
    )
    const nextAssignments: StaffingAssignment = { ...assignments }
    const nextArrivalAssignments = { ...arrivalAssignments }

    // Remove assignments for staff who are not available before rebuilding gaps.
    programme.rows
      .filter((row) => row.day === day)
      .forEach((row) =>
        row.cells.forEach((cell) => {
          const key = cellKey(row.id, cell.group)
          const assignedId = nextAssignments[key]
          if (
            assignedId &&
            (!workingIds.has(assignedId) || sickIds.has(assignedId))
          ) {
            delete nextAssignments[key]
          }
        }),
      )

    for (const row of programme.rows.filter(
      (item) => item.day === day && item.session === '3' && item.schoolLabel,
    )) {
      const key = arrivalKey(row)
      const current = nextArrivalAssignments[key] ?? { guideIds: [] }
      const leaderAvailable =
        current.leaderId &&
        workingIds.has(current.leaderId) &&
        !sickIds.has(current.leaderId)
      const guideIds = current.guideIds.map((id) =>
        id && workingIds.has(id) && !sickIds.has(id) ? id : '',
      )
      nextArrivalAssignments[key] = {
        leaderId: leaderAvailable ? current.leaderId : undefined,
        guideIds,
      }
    }

    const workload = new Map<string, number>()
    Object.values(nextAssignments).forEach((staffId) => {
      workload.set(staffId, (workload.get(staffId) ?? 0) + 1)
    })

    const dayRows = programme.rows.filter((row) => row.day === day)

    // Fill normal qualified activity assignments first.
    for (const row of dayRows) {
      for (const cell of row.cells) {
        if (!cell.activityCode || cell.activityCode === 'Z') continue

        const key = cellKey(row.id, cell.group)
        if (nextAssignments[key]) continue

        const candidates = staff
          .filter(
            (member) =>
              workingIds.has(member.id) &&
              !sickIds.has(member.id) &&
              member.qualifications.includes(cell.activityCode),
          )
          .filter((member) => {
            return !dayRows.some(
              (otherRow) =>
                otherRow.session === row.session &&
                otherRow.cells.some(
                  (otherCell) =>
                    nextAssignments[cellKey(otherRow.id, otherCell.group)] ===
                    member.id,
                ),
            )
          })
          .sort((a, b) => {
            const priorityDifference =
              rolePriority(resolvedRole(a)) - rolePriority(resolvedRole(b))
            if (priorityDifference !== 0) return priorityDifference

            const workloadDifference =
              (workload.get(a.id) ?? 0) - (workload.get(b.id) ?? 0)
            if (workloadDifference !== 0) return workloadDifference

            return a.name.localeCompare(b.name)
          })

        const chosen = candidates[0]
        if (chosen) {
          nextAssignments[key] = chosen.id
          workload.set(chosen.id, (workload.get(chosen.id) ?? 0) + 1)
        }
      }
    }

    // Session 3: assign one School Leader, then one instructor per group where
    // possible. Only pair two groups to one instructor when staffing is short.
    const arrivalRowsForDay = dayRows.filter(
      (row) => row.session === '3' && row.schoolLabel,
    )

    for (const row of arrivalRowsForDay) {
      const key = arrivalKey(row)
      const populatedGroups = row.cells
        .filter((cell) => cell.activityCode && cell.activityCode !== 'Z')
        .map((cell) => cell.group)
      const current = nextArrivalAssignments[key] ?? { guideIds: [] }

      const busyInSession = new Set<string>()
      dayRows
        .filter((item) => item.session === row.session)
        .forEach((item) =>
          item.cells.forEach((cell) => {
            const assigned = nextAssignments[cellKey(item.id, cell.group)]
            if (assigned) busyInSession.add(assigned)
          }),
        )

      current.guideIds.filter(Boolean).forEach((id) => busyInSession.add(id))
      if (current.leaderId) busyInSession.add(current.leaderId)

      if (!current.leaderId) {
        const leader = staff
          .filter(
            (member) =>
              workingIds.has(member.id) &&
              !sickIds.has(member.id) &&
              !busyInSession.has(member.id),
          )
          .sort((a, b) => {
            const leaderRank = (member: StaffMember) => {
              const role = resolvedRole(member)
              if (role === 'teamLeader') return 0
              if (role === 'activityManager') return 1
              if (role === 'centreManager') return 2
              return 3
            }
            const rankDifference = leaderRank(a) - leaderRank(b)
            if (rankDifference !== 0) return rankDifference
            return (workload.get(a.id) ?? 0) - (workload.get(b.id) ?? 0)
          })[0]

        if (leader) {
          current.leaderId = leader.id
          busyInSession.add(leader.id)
          workload.set(leader.id, (workload.get(leader.id) ?? 0) + 1)
        }
      }

      const guideIds = Array.from(
        { length: populatedGroups.length },
        (_, index) => current.guideIds[index] ?? '',
      )
      const guideLoad = new Map<string, number>()
      guideIds.filter(Boolean).forEach((id) =>
        guideLoad.set(id, (guideLoad.get(id) ?? 0) + 1),
      )

      for (let index = 0; index < populatedGroups.length; index += 1) {
        if (guideIds[index]) continue

        let candidates = staff
          .filter(
            (member) =>
              workingIds.has(member.id) &&
              !sickIds.has(member.id) &&
              member.id !== current.leaderId &&
              !busyInSession.has(member.id),
          )
          .sort((a, b) => {
            const roleDifference =
              rolePriority(resolvedRole(a)) - rolePriority(resolvedRole(b))
            if (roleDifference !== 0) return roleDifference
            return (workload.get(a.id) ?? 0) - (workload.get(b.id) ?? 0)
          })

        // If everyone is already used, allow an arrival instructor to cover a
        // second group, but never more than two groups.
        if (!candidates.length) {
          candidates = staff
            .filter(
              (member) =>
                workingIds.has(member.id) &&
                !sickIds.has(member.id) &&
                member.id !== current.leaderId &&
                (guideLoad.get(member.id) ?? 0) === 1,
            )
            .sort((a, b) =>
              (workload.get(a.id) ?? 0) - (workload.get(b.id) ?? 0),
            )
        }

        const chosen = candidates[0]
        if (chosen) {
          guideIds[index] = chosen.id
          const previousGuideLoad = guideLoad.get(chosen.id) ?? 0
          guideLoad.set(chosen.id, previousGuideLoad + 1)
          if (previousGuideLoad === 0) busyInSession.add(chosen.id)
          workload.set(chosen.id, (workload.get(chosen.id) ?? 0) + 1)
        }
      }

      nextArrivalAssignments[key] = {
        leaderId: current.leaderId,
        guideIds,
      }
    }

    setAssignments(nextAssignments)
    setArrivalAssignments(nextArrivalAssignments)
    localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(nextAssignments))
    localStorage.setItem(
      ARRIVAL_ASSIGNMENTS_KEY,
      JSON.stringify(nextArrivalAssignments),
    )
    setImportMessage(
      `Auto-filled ${day}, including a School Leader and Session 3 group instructors. One instructor per group was used where possible; groups were paired only when staffing was short.`,
    )
  }

  function isDoubleBooked(
    staffId: string,
    targetRow: ProgrammeRow,
    targetGroup: number,
  ) {
    if (!programme) return false
    return programme.rows.some(
      (row) =>
        row.day === targetRow.day &&
        row.session === targetRow.session &&
        row.cells.some(
          (cell) =>
            !(row.id === targetRow.id && cell.group === targetGroup) &&
            assignments[cellKey(row.id, cell.group)] === staffId,
        ),
    )
  }

  function toggleSick(day: string, staffId: string) {
    const current = sicknessByDay[day] ?? []
    const nextForDay = current.includes(staffId)
      ? current.filter((id) => id !== staffId)
      : [...current, staffId]
    const next = { ...sicknessByDay, [day]: nextForDay }
    setSicknessByDay(next)
    localStorage.setItem(SICKNESS_KEY, JSON.stringify(next))

    if (!current.includes(staffId) && programme) {
      const nextAssignments = { ...assignments }
      programme.rows
        .filter((row) => row.day === day)
        .forEach((row) =>
          row.cells.forEach((cell) => {
            const key = cellKey(row.id, cell.group)
            if (nextAssignments[key] === staffId) delete nextAssignments[key]
          }),
        )
      setAssignments(nextAssignments)
      localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(nextAssignments))
    }
  }

  function toggleWorking(day: string, staffId: string) {
    const current = workingByDay[day] ?? staff.map((member) => member.id)
    const isWorking = current.includes(staffId)
    const nextForDay = isWorking
      ? current.filter((id) => id !== staffId)
      : [...current, staffId]
    const next = { ...workingByDay, [day]: nextForDay }
    setWorkingByDay(next)
    localStorage.setItem(WORKING_KEY, JSON.stringify(next))

    if (isWorking && programme) {
      const nextAssignments = { ...assignments }
      programme.rows
        .filter((row) => row.day === day)
        .forEach((row) =>
          row.cells.forEach((cell) => {
            const key = cellKey(row.id, cell.group)
            if (nextAssignments[key] === staffId) delete nextAssignments[key]
          }),
        )
      setAssignments(nextAssignments)
      localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(nextAssignments))
    }
  }

  function addActivity() {
    const code = newActivityCode.trim().toUpperCase()
    const name = newActivityName.trim()
    if (!code || !name) {
      setImportMessage('Enter an activity code and activity name.')
      return
    }
    if (activities.some((activity) => activity.code === code)) {
      setImportMessage(`${code} already exists.`)
      return
    }
    const next = [...activities, { code, name }].sort((a, b) =>
      a.code.localeCompare(b.code),
    )
    setActivities(next)
    localStorage.setItem(ACTIVITIES_KEY, JSON.stringify(next))
    setNewActivityCode('')
    setNewActivityName('')
  }

  function renameActivity(code: string, name: string) {
    const next = activities.map((activity) =>
      activity.code === code ? { ...activity, name } : activity,
    )
    setActivities(next)
    localStorage.setItem(ACTIVITIES_KEY, JSON.stringify(next))
  }

  function deleteActivity(code: string) {
    if (!window.confirm(`Delete ${code} from the activity list?`)) return
    const next = activities.filter((activity) => activity.code !== code)
    setActivities(next)
    localStorage.setItem(ACTIVITIES_KEY, JSON.stringify(next))

    const nextStaff = staff.map((member) => ({
      ...member,
      qualifications: member.qualifications.filter((item) => item !== code),
      signOffs: Object.fromEntries(
        Object.entries(member.signOffs ?? {}).filter(([item]) => item !== code),
      ),
    }))
    setStaff(nextStaff)
    localStorage.setItem(STAFF_KEY, JSON.stringify(nextStaff))
  }

  function arrivalKey(row: ProgrammeRow) {
    return `${row.day}::${row.id}`
  }

  function setArrivalLeader(row: ProgrammeRow, staffId: string) {
    const key = arrivalKey(row)
    const next = {
      ...arrivalAssignments,
      [key]: {
        ...(arrivalAssignments[key] ?? { guideIds: [] }),
        leaderId: staffId || undefined,
      },
    }
    setArrivalAssignments(next)
    localStorage.setItem(
      ARRIVAL_ASSIGNMENTS_KEY,
      JSON.stringify(next),
    )
  }

  function setArrivalGuide(
    row: ProgrammeRow,
    slotIndex: number,
    staffId: string,
  ) {
    const key = arrivalKey(row)
    const current = arrivalAssignments[key] ?? { guideIds: [] }
    const guideIds = [...current.guideIds]
    guideIds[slotIndex] = staffId
    const next = {
      ...arrivalAssignments,
      [key]: { ...current, guideIds },
    }
    setArrivalAssignments(next)
    localStorage.setItem(
      ARRIVAL_ASSIGNMENTS_KEY,
      JSON.stringify(next),
    )
  }

  function assignStaff(staffId?: string) {
    if (!selectedStaffingCell) return
    const key = cellKey(
      selectedStaffingCell.row.id,
      selectedStaffingCell.group,
    )
    const next = { ...assignments }
    if (staffId) next[key] = staffId
    else delete next[key]
    setAssignments(next)
    localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(next))
    setSelectedStaffingCell(null)
  }

  const selectedStaffingCode = selectedStaffingCell
    ? selectedStaffingCell.row.cells.find(
        (cell) => cell.group === selectedStaffingCell.group,
      )?.activityCode ?? ''
    : ''

  const eligibleStaff = selectedStaffingCell
    ? staff.filter(
        (member) =>
          (workingByDay[selectedStaffingCell.row.day] ??
            staff.map((item) => item.id)
          ).includes(member.id) &&
          member.qualifications.includes(selectedStaffingCode) &&
          !(sicknessByDay[selectedStaffingCell.row.day] ?? []).includes(member.id) &&
          !isDoubleBooked(
            member.id,
            selectedStaffingCell.row,
            selectedStaffingCell.group,
          ),
      )
    : []


  async function syncRotaToStaff(showMessage = false) {
    if (!programme) return
    setCloudSyncStatus('syncing')

    const emailByStaffId = new Map(
      staff.map((member) => [
        member.id,
        (member.email ?? '').trim().toLowerCase(),
      ]),
    )

    const publishedRows: {
      programme_name: string
      day: string
      session: string
      activity_code: string
      activity_name: string
      group_numbers: number[]
      duty_type: string
      staff_email: string
      staff_name: string
      school_name: string | null
    }[] = []

    for (const row of programme.rows) {
      for (const cell of row.cells) {
        const staffId = assignments[cellKey(row.id, cell.group)]
        if (!staffId || !cell.activityCode || cell.activityCode === 'Z') {
          continue
        }

        const member = staff.find((item) => item.id === staffId)
        const email = emailByStaffId.get(staffId)
        if (!member || !email) continue

        publishedRows.push({
          programme_name: programme.title,
          day: row.day,
          session: row.session,
          activity_code: cell.activityCode,
          activity_name: activityName(cell.activityCode),
          group_numbers: [cell.group],
          duty_type: 'activity',
          staff_email: email,
          staff_name: member.name,
          school_name: row.schoolLabel ?? null,
        })
      }
    }

    for (const [key, arrival] of Object.entries(arrivalAssignments)) {
      const [day, rowId] = key.split('::')
      const row = programme.rows.find((item) => item.id === rowId)
      if (!row) continue

      if (arrival.leaderId) {
        const leader = staff.find((item) => item.id === arrival.leaderId)
        const email = emailByStaffId.get(arrival.leaderId)
        if (leader && email) {
          publishedRows.push({
            programme_name: programme.title,
            day,
            session: row.session,
            activity_code: 'ARRIVAL',
            activity_name: 'School Leader – documents and welcome',
            group_numbers: row.cells
              .filter((cell) => cell.activityCode && cell.activityCode !== 'Z')
              .map((cell) => cell.group),
            duty_type: 'arrival_leader',
            staff_email: email,
            staff_name: leader.name,
            school_name: row.schoolLabel ?? null,
          })
        }
      }

      arrival.guideIds.forEach((staffId, index) => {
        if (!staffId) return
        const guide = staff.find((item) => item.id === staffId)
        const email = emailByStaffId.get(staffId)
        if (!guide || !email) return

        publishedRows.push({
          programme_name: programme.title,
          day,
          session: row.session,
          activity_code: 'ARRIVAL',
          activity_name: 'Accommodation and fire alarm instructor',
          group_numbers: [row.cells.filter((cell) => cell.activityCode && cell.activityCode !== 'Z')[index]?.group ?? index + 1],
          duty_type: 'arrival_instructor',
          staff_email: email,
          staff_name: guide.name,
          school_name: row.schoolLabel ?? null,
        })
      })
    }

    const missingEmails = new Set(
      Object.values(assignments)
        .filter(Boolean)
        .filter((staffId) => !emailByStaffId.get(staffId))
        .map((staffId) => staff.find((item) => item.id === staffId)?.name)
        .filter(Boolean),
    )

    const { error: deleteError } = await supabase
      .from('rota_assignments')
      .delete()
      .neq('id', '00000000-0000-0000-0000-000000000000')

    if (deleteError) {
      setCloudSyncStatus('error')
      if (showMessage) setImportMessage(`Cloud sync failed: ${deleteError.message}`)
      return
    }

    if (publishedRows.length) {
      const { error: insertError } = await supabase
        .from('rota_assignments')
        .insert(publishedRows)

      if (insertError) {
        setCloudSyncStatus('error')
        if (showMessage) setImportMessage(`Cloud sync failed: ${insertError.message}`)
        return
      }
    }

    setCloudSyncStatus('synced')
    if (showMessage) {
      setImportMessage(
        `Synced ${publishedRows.length} duties to staff accounts.${
          missingEmails.size
            ? ` Add emails for: ${Array.from(missingEmails).join(', ')}.`
            : ''
        }`,
      )
    }
  }

  useEffect(() => {
    if (!programme || !accountEmail) return
    if (syncTimeoutRef.current) clearTimeout(syncTimeoutRef.current)
    syncTimeoutRef.current = setTimeout(() => {
      syncRotaToStaff(false)
    }, 900)
    return () => {
      if (syncTimeoutRef.current) clearTimeout(syncTimeoutRef.current)
    }
  }, [programme, assignments, arrivalAssignments, staff, accountEmail])

  function exportDailyStaffing(day: string) {
    if (!programme || !day) return

    const rowsForDay = programme.rows
      .filter((row) => row.day === day)
      .flatMap((row) =>
        row.cells
          .filter((cell) => cell.activityCode && cell.activityCode !== 'Z')
          .map((cell) => {
            const assignedId = assignments[cellKey(row.id, cell.group)]
            const member = staff.find((item) => item.id === assignedId)
            return {
              Day: row.day,
              Session: row.session,
              Group: cell.group,
              'Activity Code': cell.activityCode,
              Activity: activityName(cell.activityCode),
              Instructor: member?.name ?? 'UNASSIGNED',
              Role: member ? roleLabel(resolvedRole(member)) : '',
              Status: member ? 'Staffed' : 'Needs instructor',
            }
          }),
      )
      .sort((a, b) => {
        const sessionDifference =
          Number(a.Session) - Number(b.Session)
        if (sessionDifference !== 0) return sessionDifference
        return Number(a.Group) - Number(b.Group)
      })

    const byInstructor = staff
      .map((member) => {
        const jobs = rowsForDay.filter((row) => row.Instructor === member.name)
        return jobs.map((job) => ({
          Instructor: member.name,
          Role: roleLabel(resolvedRole(member)),
          Session: job.Session,
          Group: job.Group,
          Activity: job.Activity,
          'Activity Code': job['Activity Code'],
        }))
      })
      .flat()

    const workbook = XLSX.utils.book_new()
    const staffingSheet = XLSX.utils.json_to_sheet(rowsForDay)
    staffingSheet['!cols'] = [
      { wch: 10 },
      { wch: 10 },
      { wch: 8 },
      { wch: 16 },
      { wch: 26 },
      { wch: 22 },
      { wch: 20 },
      { wch: 18 },
    ]
    XLSX.utils.book_append_sheet(workbook, staffingSheet, `${day} Staffing`)

    const instructorSheet = XLSX.utils.json_to_sheet(byInstructor)
    instructorSheet['!cols'] = [
      { wch: 22 },
      { wch: 20 },
      { wch: 10 },
      { wch: 8 },
      { wch: 26 },
      { wch: 16 },
    ]
    XLSX.utils.book_append_sheet(workbook, instructorSheet, 'By Instructor')

    const sessions = Array.from(
      new Set(rowsForDay.map((row) => String(row.Session))),
    ).sort((a, b) => Number(a) - Number(b))

    const wallRows = staff
      .map((member) => {
        const wallRow: Record<string, string> = {
          Instructor: member.name,
          Role: roleLabel(resolvedRole(member)),
        }
        for (const session of sessions) {
          const duties = rowsForDay.filter(
            (item) =>
              String(item.Session) === session &&
              item.Instructor === member.name,
          )
          wallRow[`Session ${session}`] = duties.length
            ? duties
                .map((duty) => `${duty['Activity Code']} · G${duty.Group}`)
                .join(' / ')
            : 'FREE'
        }
        return wallRow
      })
      .filter((row) =>
        sessions.some((session) => row[`Session ${session}`] !== 'FREE'),
      )

    const wallSheet = XLSX.utils.json_to_sheet(wallRows)
    wallSheet['!cols'] = [
      { wch: 22 },
      { wch: 20 },
      ...sessions.map(() => ({ wch: 22 })),
    ]
    XLSX.utils.book_append_sheet(
      workbook,
      wallSheet,
      'Staff Room Wall Rota',
    )

    XLSX.writeFile(workbook, `${day}-daily-staffing.xlsx`)
  }

  const filteredStaffingCells = populatedCells.filter(({ row, cell }) => {
    if (activeStaffingDay && row.day !== activeStaffingDay) return false
    const staffId = assignments[cellKey(row.id, cell.group)]
    const staffName = staff.find((member) => member.id === staffId)?.name ?? ''
    return `${row.day} ${row.session} group ${cell.group} ${cell.activityCode} ${activityName(cell.activityCode)} ${staffName}`
      .toLowerCase()
      .includes(query.toLowerCase())
  })

  const arrivalRows =
    programme?.rows.filter(
      (row) => row.schoolLabel && row.session === '3',
    ) ?? []

  const selectedDayDemand = sessionDemandForDay(activeStaffingDay)
  const selectedDayBusiest = busiestSessionForDay(activeStaffingDay)
  const programmeBusiest = busiestSessionAcrossProgramme()

  const criticalManagersOnSession = staff.filter((member) => {
    const role = resolvedRole(member)
    if (role !== 'activityManager' && role !== 'centreManager') return false
    return Object.entries(assignments).some(([key, staffId]) => {
      if (staffId !== member.id) return false
      const [rowId] = key.split('::')
      const row = programme?.rows.find((item) => item.id === rowId)
      return row?.day === activeStaffingDay
    })
  })

  return (
    <div className="app-shell">
      <input
        ref={fileInputRef}
        className="hidden-input"
        type="file"
        accept=".xlsx,.xls"
        onChange={importExcel}
      />

      <header className="topbar">
        <div>
          <p className="eyebrow">Norfolk Lakes</p>
          <h1>Adventure Centre Manager</h1>
          <small className="account-email">{accountEmail}</small>
        </div>
        <div className="account-actions">
          <div className={`sync-status sync-${cloudSyncStatus}`}>
            <span className="sync-dot" />
            {cloudSyncStatus === 'syncing'
              ? 'Syncing rota…'
              : cloudSyncStatus === 'error'
                ? 'Sync problem'
                : cloudSyncStatus === 'synced'
                  ? 'Staff rota live'
                  : 'Cloud connected'}
          </div>
          <button
            className="upload-top"
            onClick={() => fileInputRef.current?.click()}
          >
            <Upload size={17} />
            Upload programme
          </button>
          <button className="sign-out-top" onClick={onSignOut}>
            <LogOut size={17} />
            Sign out
          </button>
        </div>
      </header>

      <Nav page={page} setPage={setPage} />

      <main className="page-content">
        {importMessage && (
          <div className="import-message">{importMessage}</div>
        )}

        {page === 'dashboard' && (
          <>
            <section className="hero">
              <p className="eyebrow">Programme control</p>
              <h2>{programme?.title ?? 'Upload today’s programme'}</h2>
              <p>
                Import the Excel programme, make same-day changes and generate
                qualified staffing from the same grid.
              </p>
              <button
                className="hero-upload"
                onClick={() => fileInputRef.current?.click()}
              >
                <FileSpreadsheet size={20} />
                {programme ? 'Replace programme' : 'Upload Excel programme'}
              </button>
            </section>

            <section className="stats-grid">
              <Stat
                icon={<CalendarDays />}
                value={programme?.rows.length ?? 0}
                label="Programme rows"
              />
              <Stat
                icon={<Users />}
                value={populatedCells.length}
                label="Activity places"
              />
              <Stat
                icon={<Users />}
                value={programmeBusiest?.total ?? 0}
                label="Maximum staff needed"
              />
              <article className="stat-card busiest-card">
                <span><CircleAlert /></span>
                <strong>
                  {programmeBusiest
                    ? `${programmeBusiest.day} · S${programmeBusiest.session}`
                    : '—'}
                </strong>
                <small>Busiest day and session</small>
              </article>
            </section>

            <section className="action-grid">
              <Action
                icon={<FileSpreadsheet />}
                title="Programme grid"
                subtitle="View and edit the imported spreadsheet"
                onClick={() => setPage('programme')}
              />
              <Action
                icon={<Users />}
                title="Daily staffing"
                subtitle="Assign qualified staff to programme cells"
                onClick={() => setPage('staffing')}
              />
              <Action
                icon={<Users />}
                title="Staff management"
                subtitle={`${staff.length} staff members imported from the sign-off sheet`}
                onClick={() => setPage('staff')}
              />
              <Action
                icon={<ShieldCheck />}
                title="Sign-offs"
                subtitle="Manage staff activity permissions"
                onClick={() => setPage('signoffs')}
              />
              <Action
                icon={<History />}
                title="Previous versions"
                subtitle={`${history.length} saved programme version${history.length === 1 ? '' : 's'}`}
                onClick={() => setPage('programme')}
              />
            </section>
          </>
        )}

        {page === 'programme' && (
          <Panel title="Programme grid" onBack={() => setPage('dashboard')}>
            <div className="programme-toolbar">
              <button
                className="primary"
                onClick={() => fileInputRef.current?.click()}
              >
                <Upload size={18} />
                {programme ? 'Upload changed programme' : 'Upload programme'}
              </button>
              {programme && (
                <div className="programme-details">
                  <strong>{programme.sourceFileName}</strong>
                  <span>
                    Sheet: {programme.sheetName} · Imported{' '}
                    {new Date(programme.importedAt).toLocaleString()}
                  </span>
                </div>
              )}
            </div>

            {!programme ? (
              <EmptyProgramme onUpload={() => fileInputRef.current?.click()} />
            ) : (
              <>
                <ProgrammeGrid
                  programme={programme}
                  activities={activities}
                  onSelect={(row, group) => setSelectedCell({ row, group })}
                />

                {history.length > 0 && (
                  <section className="history-section">
                    <h3>Previous uploads</h3>
                    <div className="history-list">
                      {history.map((version, index) => (
                        <article key={`${version.importedAt}-${index}`}>
                          <div>
                            <strong>{version.sourceFileName}</strong>
                            <small>
                              {new Date(version.importedAt).toLocaleString()}
                            </small>
                          </div>
                          <button onClick={() => restoreVersion(version)}>
                            Restore
                          </button>
                        </article>
                      ))}
                    </div>
                  </section>
                )}
              </>
            )}
          </Panel>
        )}

        {page === 'staffing' && (
          <Panel title="Daily staffing" onBack={() => setPage('dashboard')}>
            {!programme ? (
              <EmptyProgramme onUpload={() => fileInputRef.current?.click()} />
            ) : (
              <>
                <div className="staffing-controls">
                  <div className="day-tabs" role="tablist" aria-label="Staffing day">
                    {programmeDays.map((day) => (
                      <button
                        key={day}
                        className={activeStaffingDay === day ? 'active' : ''}
                        onClick={() => {
                          setSelectedStaffingDay(day)
                          setShowSickPanel(false)
                        }}
                      >
                        {day}
                      </button>
                    ))}
                  </div>

                  <div className="staffing-actions">
                    <button
                      className="secondary-action"
                      onClick={() =>
                        setShowWorkingPanel((current) => !current)
                      }
                    >
                      <Users size={17} />
                      Staff in today
                      <span className="count-badge neutral">
                        {(
                          workingByDay[activeStaffingDay] ??
                          staff.map((member) => member.id)
                        ).length}
                      </span>
                    </button>

                    <button
                      className="secondary-action"
                      onClick={() => setShowSickPanel((current) => !current)}
                    >
                      <Users size={17} />
                      Sick staff
                      {(sicknessByDay[activeStaffingDay] ?? []).length > 0 && (
                        <span className="count-badge">
                          {(sicknessByDay[activeStaffingDay] ?? []).length}
                        </span>
                      )}
                    </button>

                    <button
                      className="clear-staffing-button"
                      onClick={() => clearDayStaffing(activeStaffingDay)}
                    >
                      <X size={17} />
                      Clear day
                    </button>

                    <button
                      className="auto-fill-button"
                      onClick={() => autoFillStaffing(activeStaffingDay)}
                    >
                      <WandSparkles size={17} />
                      Auto-fill staff
                    </button>
                    <button
                      className="print-button"
                      onClick={() => exportDailyStaffing(activeStaffingDay)}
                    >
                      <Printer size={17} />
                      Print / Excel
                    </button>
                  </div>
                </div>

                {showWorkingPanel && (
                  <section className="sickness-panel compact working-panel">
                    <div className="sickness-heading">
                      <div>
                        <p className="eyebrow">Daily availability</p>
                        <h3>Who is working on {activeStaffingDay}?</h3>
                      </div>
                      <button
                        className="icon-button small"
                        onClick={() => setShowWorkingPanel(false)}
                      >
                        <X size={17} />
                      </button>
                    </div>
                    <p>
                      Auto-fill and manual assignment only use people selected
                      as working today.
                    </p>
                    <div className="sickness-chips">
                      {staff.map((member) => {
                        const working = (
                          workingByDay[activeStaffingDay] ??
                          staff.map((item) => item.id)
                        ).includes(member.id)
                        return (
                          <button
                            key={member.id}
                            className={working ? 'working active' : 'working'}
                            onClick={() =>
                              toggleWorking(activeStaffingDay, member.id)
                            }
                          >
                            {member.name}
                            <span>{working ? 'Working' : 'Not in'}</span>
                          </button>
                        )
                      })}
                    </div>
                  </section>
                )}

                {showSickPanel && (
                  <section className="sickness-panel compact">
                    <div className="sickness-heading">
                      <div>
                        <p className="eyebrow">Unavailable today</p>
                        <h3>Who is sick on {activeStaffingDay}?</h3>
                      </div>
                      <button
                        className="icon-button small"
                        onClick={() => setShowSickPanel(false)}
                      >
                        <X size={17} />
                      </button>
                    </div>
                    <p>
                      Sick staff are removed from this day and cannot be selected
                      or included by auto-fill.
                    </p>
                    <div className="sickness-chips">
                      {staff.map((member) => {
                        const sick = (
                          sicknessByDay[activeStaffingDay] ?? []
                        ).includes(member.id)
                        return (
                          <button
                            key={member.id}
                            className={sick ? 'sick active' : 'sick'}
                            onClick={() =>
                              toggleSick(activeStaffingDay, member.id)
                            }
                          >
                            {member.name}
                            <span>{sick ? 'Sick' : 'Available'}</span>
                          </button>
                        )
                      })}
                    </div>
                  </section>
                )}

                {criticalManagersOnSession.length > 0 && (
                  <section className="management-warning">
                    <CircleAlert size={22} />
                    <div>
                      <strong>Centre management cover warning</strong>
                      <p>
                        {criticalManagersOnSession
                          .map(
                            (member) =>
                              `${member.name} (${roleLabel(
                                resolvedRole(member),
                              )})`,
                          )
                          .join(', ')}{' '}
                        {criticalManagersOnSession.length === 1
                          ? 'is'
                          : 'are'} assigned to sessions on {activeStaffingDay}.
                        Confirm that someone suitable remains free to run the
                        centre.
                      </p>
                    </div>
                  </section>
                )}

                {arrivalRows
                  .filter((row) => row.day === activeStaffingDay)
                  .map((row) => {
                    const populatedGroups = row.cells.filter(
                      (cell) => cell.activityCode && cell.activityCode !== 'Z',
                    )
                    const groupCount = populatedGroups.length
                    const guideSlots = Math.max(1, groupCount)
                    const key = arrivalKey(row)
                    const assignment =
                      arrivalAssignments[key] ?? { guideIds: [] }
                    const availableToday =
                      workingByDay[activeStaffingDay] ??
                      staff.map((member) => member.id)
                    const sickToday =
                      sicknessByDay[activeStaffingDay] ?? []

                    const leaderOptions = staff
                      .filter(
                        (member) =>
                          availableToday.includes(member.id) &&
                          !sickToday.includes(member.id) &&
                          !assignment.guideIds.includes(member.id),
                      )
                      .sort((a, b) => {
                        const leaderRank = (member: StaffMember) => {
                          const role = resolvedRole(member)
                          if (role === 'teamLeader') return 0
                          if (role === 'activityManager') return 1
                          if (role === 'centreManager') return 2
                          return 3
                        }
                        return leaderRank(a) - leaderRank(b)
                      })

                    const guideOptions = staff.filter(
                      (member) =>
                        availableToday.includes(member.id) &&
                        !sickToday.includes(member.id) &&
                        member.id !== assignment.leaderId &&
                        (arrivalStaffGroup === 'all' ||
                          resolvedRole(member) === arrivalStaffGroup),
                    )

                    return (
                      <section className="arrival-card" key={row.id}>
                        <div className="arrival-card-heading">
                          <div>
                            <p className="eyebrow">School welcome and accommodation · Session 3</p>
                            <h3>{row.schoolLabel}</h3>
                            <p>
                              {groupCount} group{groupCount === 1 ? '' : 's'} · one School Leader plus up to {guideSlots} group instructor{guideSlots === 1 ? '' : 's'}
                            </p>
                          </div>
                        </div>

                        <div className="arrival-staff-filter">
                          <label>
                            Instructor staff group
                            <select
                              value={arrivalStaffGroup}
                              onChange={(event) =>
                                setArrivalStaffGroup(
                                  event.target.value as
                                    | 'all'
                                    | StaffRole,
                                )
                              }
                            >
                              <option value="all">All working staff</option>
                              <option value="staff">Staff</option>
                              <option value="teamLeader">Team leaders</option>
                              <option value="activityManager">
                                Activities managers
                              </option>
                              <option value="centreManager">
                                Centre managers
                              </option>
                            </select>
                          </label>
                          <p>
                            The School Leader talks through documents while group instructors take children to accommodation and complete the fire alarm test. One instructor per group is preferred; one instructor may cover two groups when staffing is short.
                          </p>
                        </div>

                        <div className="arrival-assignment-grid">
                          <label>
                            School Leader
                            <select
                              value={assignment.leaderId ?? ''}
                              onChange={(event) =>
                                setArrivalLeader(row, event.target.value)
                              }
                            >
                              <option value="">Select School Leader</option>
                              {leaderOptions.map((member) => (
                                <option key={member.id} value={member.id}>
                                  {member.name}
                                </option>
                              ))}
                            </select>
                          </label>

                          {Array.from({ length: guideSlots }, (_, index) => {
                            const groupNumber = index + 1
                            return (
                              <label key={index}>
                                Instructor for Group {groupNumber}
                                <select
                                  value={assignment.guideIds[index] ?? ''}
                                  onChange={(event) =>
                                    setArrivalGuide(
                                      row,
                                      index,
                                      event.target.value,
                                    )
                                  }
                                >
                                  <option value="">Select instructor</option>
                                  {guideOptions
                                    .filter((member) => {
                                      const existingCount =
                                        assignment.guideIds.filter(
                                          (id, guideIndex) =>
                                            guideIndex !== index &&
                                            id === member.id,
                                        ).length
                                      return existingCount < 2
                                    })
                                    .map((member) => (
                                      <option
                                        key={member.id}
                                        value={member.id}
                                      >
                                        {member.name}
                                      </option>
                                    ))}
                                </select>
                              </label>
                            )
                          })}
                        </div>
                        <p className="arrival-note">
                          Preferred staffing is one instructor per group. The same instructor can be selected for a maximum of two groups when staffing is short.
                        </p>
                      </section>
                    )
                  })}

                <div className="search-box">
                  <Search size={18} />
                  <input
                    value={query}
                    onChange={(event) => setQuery(event.target.value)}
                    placeholder="Search group, activity or instructor"
                  />
                </div>

                <section className="session-demand-summary">
                  <div className="busiest-session-card">
                    <p className="eyebrow">Maximum staff needed</p>
                    <strong>{selectedDayBusiest?.total ?? 0}</strong>
                    <span>
                      {selectedDayBusiest
                        ? `${activeStaffingDay} · Session ${selectedDayBusiest.session}`
                        : activeStaffingDay}
                    </span>
                  </div>

                  <div className="session-demand-list">
                    {selectedDayDemand.map((item) => (
                      <article
                        key={item.session}
                        className={
                          selectedDayBusiest?.session === item.session
                            ? 'active'
                            : ''
                        }
                      >
                        <span>Session {item.session}</span>
                        <strong>{item.total} staff</strong>
                      </article>
                    ))}
                  </div>
                </section>

                <div className="staffing-grid">
                  {filteredStaffingCells.map(({ row, cell }) => {
                    const key = cellKey(row.id, cell.group)
                    const assignedStaff = staff.find(
                      (member) => member.id === assignments[key],
                    )
                    return (
                      <article
                        className={`staffing-card ${assignedStaff ? 'ready' : 'needs'}`}
                        key={key}
                      >
                        <div className="staffing-card-top">
                          <span>
                            {row.day} · Session {row.session}
                          </span>
                          <span>
                            {assignedStaff ? 'Ready' : 'Needs instructor'}
                          </span>
                        </div>
                        <h3>{activityName(cell.activityCode)}</h3>
                        <p>
                          Group {cell.group} · {cell.activityCode}
                        </p>
                        <div className="assignment-row">
                          <div>
                            <small>Instructor</small>
                            <strong>
                              {assignedStaff?.name ?? 'Not assigned'}
                            </strong>
                          </div>
                          <button
                            onClick={() =>
                              setSelectedStaffingCell({
                                row,
                                group: cell.group,
                              })
                            }
                          >
                            {assignedStaff ? 'Change' : 'Assign'}
                          </button>
                        </div>
                      </article>
                    )
                  })}
                </div>
              </>
            )}
          </Panel>
        )}

        {page === 'staff' && (
          <Panel title="Staff management" onBack={() => setPage('dashboard')}>
            <div className="staff-page-toolbar">
              <div>
                <p>
                  Add new instructors, set team leaders and review how many
                  activities each person is signed off to run.
                </p>
              </div>
              <button
                className="primary"
                onClick={() => setShowAddStaff((current) => !current)}
              >
                <Plus size={18} />
                Add new staff member
              </button>
            </div>

            {showAddStaff && (
              <section className="add-staff-panel">
                <div className="add-staff-heading">
                  <div>
                    <p className="eyebrow">Manager controls</p>
                    <h3>Add staff member</h3>
                  </div>
                  <button
                    className="icon-button small"
                    onClick={() => setShowAddStaff(false)}
                  >
                    <X size={17} />
                  </button>
                </div>

                <label>Staff member’s name</label>
                <input
                  value={newStaffName}
                  onChange={(event) => setNewStaffName(event.target.value)}
                  placeholder="Full name"
                />

                <label>Operational role</label>
                <div className="role-selector">
                  {(
                    [
                      ['staff', 'Staff'],
                      ['teamLeader', 'Team leader'],
                      ['activityManager', 'Activities manager'],
                      ['centreManager', 'Centre manager'],
                    ] as [StaffRole, string][]
                  ).map(([role, label]) => (
                    <button
                      key={role}
                      className={newStaffRole === role ? 'active' : ''}
                      onClick={() => setNewStaffRole(role)}
                    >
                      {label}
                    </button>
                  ))}
                </div>

                <h4>Initial activity sign-offs</h4>
                <div className="new-staff-activity-grid">
                  {activities.map((activity) => (
                    <button
                      key={activity.code}
                      className={
                        newStaffQualifications.includes(activity.code)
                          ? 'chip active'
                          : 'chip'
                      }
                      title={activity.name}
                      onClick={() =>
                        toggleNewStaffQualification(activity.code)
                      }
                    >
                      {activity.code}
                    </button>
                  ))}
                </div>

                <button className="primary save-staff" onClick={addStaffMember}>
                  Add staff member
                </button>
              </section>
            )}

            <div className="staff-management-list">
              {staff.map((member) => (
                <article className="staff-management-card" key={member.id}>
                  <div>
                    <div className="staff-name-line">
                      <h3>{member.name}</h3>
                      <span className={`role-label role-${resolvedRole(member)}`}>
                        {roleLabel(resolvedRole(member))}
                      </span>
                    </div>
                    <p>
                      {member.qualifications.length} activity sign-off
                      {member.qualifications.length === 1 ? '' : 's'}
                    </p>
                    <label className="staff-email-field">
                      Login email
                      <input
                        type="email"
                        value={member.email ?? ''}
                        placeholder="name@example.com"
                        onChange={(event) =>
                          updateStaffEmail(member.id, event.target.value)
                        }
                      />
                    </label>
                  </div>

                  <div className="staff-card-actions">
                    <select
                      className="role-select"
                      value={resolvedRole(member)}
                      onChange={(event) =>
                        setStaffRole(
                          member.id,
                          event.target.value as StaffRole,
                        )
                      }
                    >
                      <option value="staff">Staff</option>
                      <option value="teamLeader">Team leader</option>
                      <option value="activityManager">
                        Activities manager
                      </option>
                      <option value="centreManager">Centre manager</option>
                    </select>
                    <button
                      className="delete-staff"
                      onClick={() => deleteStaffMember(member.id)}
                    >
                      <Trash2 size={16} />
                      Remove
                    </button>
                  </div>
                </article>
              ))}
            </div>
          </Panel>
        )}

        {page === 'signoffs' && (
          <Panel title="Staff sign-offs" onBack={() => setPage('dashboard')}>
            <div className="signoff-toolbar">
              <p>
                {staff.length} staff members loaded. Edit sign-offs and manage
                the centre’s activity list.
              </p>
              <button
                className="primary"
                onClick={() =>
                  setShowActivityManager((current) => !current)
                }
              >
                Manage activities
              </button>
            </div>

            {showActivityManager && (
              <section className="activity-manager-panel">
                <div className="add-activity-row">
                  <input
                    value={newActivityCode}
                    onChange={(event) =>
                      setNewActivityCode(event.target.value)
                    }
                    placeholder="Code, e.g. ZIP"
                  />
                  <input
                    value={newActivityName}
                    onChange={(event) =>
                      setNewActivityName(event.target.value)
                    }
                    placeholder="Activity name"
                  />
                  <button className="primary" onClick={addActivity}>
                    Add activity
                  </button>
                </div>

                <div className="activity-management-list">
                  {activities.map((activity) => (
                    <article key={activity.code}>
                      <strong>{activity.code}</strong>
                      <input
                        value={activity.name}
                        onChange={(event) =>
                          renameActivity(
                            activity.code,
                            event.target.value,
                          )
                        }
                      />
                      <button
                        className="delete-staff"
                        onClick={() => deleteActivity(activity.code)}
                      >
                        Delete
                      </button>
                    </article>
                  ))}
                </div>
              </section>
            )}

            <div className="signoff-list">
              {staff.map((member) => (
                <article className="signoff-card" key={member.id}>
                  <div className="signoff-card-heading">
                    <div>
                      <h3>{member.name}</h3>
                      <span className={`role-label role-${resolvedRole(member)}`}>
                        {roleLabel(resolvedRole(member))}
                      </span>
                    </div>
                    <div className="signoff-person-actions">
                      <select
                        className="role-select"
                        value={resolvedRole(member)}
                        onChange={(event) =>
                          setStaffRole(
                            member.id,
                            event.target.value as StaffRole,
                          )
                        }
                      >
                        <option value="staff">Staff</option>
                        <option value="teamLeader">Team leader</option>
                        <option value="activityManager">
                          Activities manager
                        </option>
                        <option value="centreManager">Centre manager</option>
                      </select>
                      <button
                        className="select-all-button"
                        onClick={() => selectAllQualifications(member.id)}
                      >
                        Select all activities
                      </button>
                      <button
                        className="clear-all-button"
                        onClick={() => clearAllQualifications(member.id)}
                      >
                        Clear all
                      </button>
                    </div>
                  </div>
                  <div className="chip-grid">
                    {activities.map((activity) => (
                      <button
                        key={activity.code}
                        className={
                          member.qualifications.includes(activity.code)
                            ? 'chip active'
                            : 'chip'
                        }
                        title={activity.name}
                        onClick={() =>
                          toggleQualification(member.id, activity.code)
                        }
                      >
                        <span>{activity.code}</span>
                        {member.signOffs?.[activity.code] && (
                          <small>{member.signOffs[activity.code]}</small>
                        )}
                      </button>
                    ))}
                  </div>
                </article>
              ))}
            </div>
          </Panel>
        )}
      </main>


      {selectedCell && (
        <div className="modal-backdrop">
          <section className="modal">
            <div className="modal-head">
              <div>
                <p className="eyebrow">Change activity</p>
                <h2>
                  {selectedCell.row.day} · Session{' '}
                  {selectedCell.row.session} · Group {selectedCell.group}
                </h2>
              </div>
              <button
                className="icon-button"
                onClick={() => setSelectedCell(null)}
              >
                <X />
              </button>
            </div>
            <div className="activity-picker">
              <button
                onClick={() =>
                  updateActivity(
                    selectedCell.row.id,
                    selectedCell.group,
                    '',
                  )
                }
              >
                Clear cell
              </button>
              {activities.map((activity) => (
                <button
                  key={activity.code}
                  onClick={() =>
                    updateActivity(
                      selectedCell.row.id,
                      selectedCell.group,
                      activity.code,
                    )
                  }
                >
                  <strong>{activity.code}</strong>
                  <span>{activity.name}</span>
                </button>
              ))}
            </div>
          </section>
        </div>
      )}

      {selectedStaffingCell && (
        <div className="modal-backdrop">
          <section className="modal">
            <div className="modal-head">
              <div>
                <p className="eyebrow">Assign instructor</p>
                <h2>{activityName(selectedStaffingCode)}</h2>
                <p>
                  {selectedStaffingCell.row.day} · Session{' '}
                  {selectedStaffingCell.row.session} · Group{' '}
                  {selectedStaffingCell.group}
                </p>
              </div>
              <button
                className="icon-button"
                onClick={() => setSelectedStaffingCell(null)}
              >
                <X />
              </button>
            </div>

            <div className="picker-list">
              {eligibleStaff.length ? (
                eligibleStaff.map((member) => (
                  <button
                    key={member.id}
                    onClick={() => assignStaff(member.id)}
                  >
                    <span>
                      <strong>{member.name}</strong>
                      <small>Qualified and available</small>
                    </span>
                    <span>Assign</span>
                  </button>
                ))
              ) : (
                <p>
                  No qualified, available instructor is currently listed for
                  this activity and session.
                </p>
              )}
            </div>

            <button
              className="remove-button"
              onClick={() => assignStaff(undefined)}
            >
              Remove instructor
            </button>
          </section>
        </div>
      )}
    </div>
  )
}

function ProgrammeGrid({
  programme,
  activities,
  onSelect,
}: {
  programme: ProgrammeImport
  activities: Activity[]
  onSelect: (row: ProgrammeRow, group: number) => void
}) {
  return (
    <div className="programme-scroll">
      <table className="programme-table">
        <thead>
          <tr>
            <th className="sticky-day">Day</th>
            <th className="sticky-session">Ses</th>
            {programme.groupNumbers.map((group) => (
              <th key={group}>G{group}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {programme.rows.map((row, index) => {
            const previous = programme.rows[index - 1]
            const showDay = !previous || previous.day !== row.day
            return (
              <tr key={row.id}>
                <th className="sticky-day">
                  {showDay ? row.day : ''}
                  {row.schoolLabel && (
                    <small>{row.schoolLabel}</small>
                  )}
                </th>
                <th className="sticky-session">{row.session}</th>
                {programme.groupNumbers.map((group) => {
                  const cell = row.cells.find(
                    (item) => item.group === group,
                  )
                  const code = cell?.activityCode ?? ''
                  return (
                    <td key={group}>
                      <button
                        className={`programme-cell code-${code.toLowerCase()}`}
                        onClick={() => onSelect(row, group)}
                        title={code ? activityNameFromList(activities, code) : 'Empty'}
                      >
                        {code || '—'}
                      </button>
                    </td>
                  )
                })}
              </tr>
            )
          })}
        </tbody>
      </table>
    </div>
  )
}

function EmptyProgramme({ onUpload }: { onUpload: () => void }) {
  return (
    <section className="empty-state">
      <FileSpreadsheet size={42} />
      <h3>No programme uploaded</h3>
      <p>Select one of your existing annual Excel programme files.</p>
      <button className="primary" onClick={onUpload}>
        <Upload size={18} />
        Upload Excel programme
      </button>
    </section>
  )
}

function Panel({
  title,
  onBack,
  children,
}: {
  title: string
  onBack: () => void
  children: React.ReactNode
}) {
  return (
    <section className="panel">
      <button className="back" onClick={onBack}>
        <ChevronLeft size={18} />
        Back
      </button>
      <h2>{title}</h2>
      {children}
    </section>
  )
}

function Stat({
  icon,
  value,
  label,
  warning = false,
}: {
  icon: React.ReactNode
  value: number
  label: string
  warning?: boolean
}) {
  return (
    <article className={`stat-card ${warning ? 'warning' : ''}`}>
      <span>{icon}</span>
      <strong>{value}</strong>
      <small>{label}</small>
    </article>
  )
}

function Action({
  icon,
  title,
  subtitle,
  onClick,
}: {
  icon: React.ReactNode
  title: string
  subtitle: string
  onClick: () => void
}) {

  return (
    <button className="action-card" onClick={onClick}>
      <span>{icon}</span>
      <div>
        <strong>{title}</strong>
        <small>{subtitle}</small>
      </div>
    </button>
  )
}

export default ManagerApp
