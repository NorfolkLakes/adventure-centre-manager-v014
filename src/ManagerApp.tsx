import { ChangeEvent, useEffect, useMemo, useRef, useState } from 'react'
import {
  Building2,
  CalendarDays,
  CalendarRange,
  ChevronRight,
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
  Bot,
  CloudSun,
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


type MySessionDuty = {
  id: string
  programme_name: string
  day: string
  session: string
  activity_name: string
  group_numbers: number[]
  duty_type: string
  school_name: string | null
  building_name: string | null
  party_leader_name: string | null
}

function readJson<T>(key: string, fallback: T): T {
  try {
    const value = localStorage.getItem(key)
    return value ? JSON.parse(value) : fallback
  } catch {
    return fallback
  }
}

function shuffled<T>(items: T[]): T[] {
  const next = [...items]
  for (let index = next.length - 1; index > 0; index -= 1) {
    const randomIndex = Math.floor(Math.random() * (index + 1))
    ;[next[index], next[randomIndex]] = [next[randomIndex], next[index]]
  }
  return next
}

const accommodationNames = ['Kingfisher', 'Swan', 'Grebe', 'Bittern', 'Mallard', 'Teal']

function accommodationName(buildingNumber: number) {
  return accommodationNames[buildingNumber - 1] ?? `Building ${buildingNumber}`
}

function normaliseActivityText(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, ' ')
}

function normaliseText(value: unknown) {
  return String(value ?? '').trim()
}

function normaliseIdentity(value: unknown) {
  return normaliseText(value).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim()
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

const ARRIVAL_DAY_ALIASES = new Set(['MON', 'MONDAY', 'WED', 'WEDNESDAY', 'FRI', 'FRIDAY'])
const PROGRAMME_ACTIVITY_VALUES = new Set(
  startingActivities.flatMap((activity) => [
    activity.code.trim().toUpperCase(),
    activity.name.trim().toUpperCase(),
  ]),
)

function normalisedProgrammeValue(value: string) {
  return value.replace(/\s+/g, ' ').trim().toUpperCase()
}

function isArrivalDay(day: string) {
  return ARRIVAL_DAY_ALIASES.has(normalisedProgrammeValue(day))
}

function isKnownProgrammeActivity(value: string) {
  const normalised = normalisedProgrammeValue(value)
  return Boolean(normalised) && PROGRAMME_ACTIVITY_VALUES.has(normalised)
}

function looksLikeSchoolName(value: string) {
  const normalised = normalisedProgrammeValue(value)
  if (!normalised || normalised === 'Z' || isKnownProgrammeActivity(normalised)) {
    return false
  }

  return /[A-Z]/i.test(normalised) && normalised.length >= 4
}

type ArrivalSchoolSegment = {
  schoolName: string
  cells: ProgrammeRow['cells']
}

function arrivalSchoolSegments(row: ProgrammeRow): ArrivalSchoolSegment[] {
  if (row.session !== '3' || !isArrivalDay(row.day)) {
    return []
  }

  const sortedCells = [...row.cells]
    .filter((cell) => cell.group >= 1 && cell.group <= 30)
    .sort((a, b) => a.group - b.group)

  const segments: ArrivalSchoolSegment[] = []
  let current: ArrivalSchoolSegment | null = null

  for (const cell of sortedCells) {
    const value = cell.activityCode.trim()

    if (looksLikeSchoolName(value)) {
      current = {
        schoolName: value.replace(/\s+/g, ' ').trim(),
        cells: [{ ...cell, activityCode: value }],
      }
      segments.push(current)
      continue
    }

    if (value && value.toUpperCase() !== 'Z' && isKnownProgrammeActivity(value)) {
      current = null
      continue
    }

    if (current && (!value || value.toUpperCase() === 'Z')) {
      current.cells.push({ ...cell, activityCode: current.schoolName })
    }
  }

  if (!segments.length && row.schoolLabel?.trim()) {
    return [{
      schoolName: row.schoolLabel.trim(),
      cells: sortedCells.filter((cell) => !cell.activityCode || cell.activityCode.toUpperCase() === 'Z'),
    }]
  }

  return segments
}

function schoolNamesInRow(row: ProgrammeRow) {
  return arrivalSchoolSegments(row).map((segment) => segment.schoolName)
}

function arrivalRowsFromProgrammeRow(row: ProgrammeRow): ProgrammeRow[] {
  return arrivalSchoolSegments(row).map((segment) => {
    const schoolKey = normalisedProgrammeValue(segment.schoolName)

    return {
      ...row,
      id: `${row.id}::arrival::${schoolKey.replace(/[^A-Z0-9]+/g, '-')}`,
      schoolLabel: segment.schoolName,
      cells: segment.cells,
    }
  })
}

function activityCellsForRow(row: ProgrammeRow) {
  const arrivalGroups = new Set(
    arrivalSchoolSegments(row).flatMap((segment) => segment.cells.map((cell) => cell.group)),
  )

  return row.cells.filter((cell) => {
    const value = cell.activityCode.trim()
    if (!value || value.toUpperCase() === 'Z') return false
    if (arrivalGroups.has(cell.group)) return false
    return isKnownProgrammeActivity(value)
  })
}

function arrivalSchoolName(row: ProgrammeRow) {
  return row.schoolLabel?.trim() ?? ''
}

function ManagerApp({
  accountEmail,
  displayName,
  onSignOut,
  accountRole = 'centreManager',
}: {
  accountEmail: string
  displayName?: string | null
  onSignOut: () => void
  accountRole?: 'centreManager' | 'activityManager' | 'teamLeader'
}) {
  const canManageHolidays = accountRole === 'centreManager' || accountRole === 'activityManager'
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
  const [weather, setWeather] = useState<{ temperature: number; wind: number; code: number } | null>(null)
  const [cloudSyncStatus, setCloudSyncStatus] = useState<
    'idle' | 'syncing' | 'synced' | 'error'
  >('idle')
  const syncTimeoutRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const [showSickPanel, setShowSickPanel] = useState(false)
  const [showWorkingPanel, setShowWorkingPanel] = useState(false)
  const [showActivityManager, setShowActivityManager] = useState(false)
  const [signoffSearch, setSignoffSearch] = useState('')
  const [newActivityCode, setNewActivityCode] = useState('')
  const [newActivityName, setNewActivityName] = useState('')
  const [arrivalStaffGroup, setArrivalStaffGroup] = useState<
    'all' | StaffRole
  >('all')
  const [showAddStaff, setShowAddStaff] = useState(false)
  const [newStaffName, setNewStaffName] = useState('')
  const [newStaffRole, setNewStaffRole] = useState<StaffRole>('staff')
  const [newStaffQualifications, setNewStaffQualifications] = useState<string[]>([])
  const [holidayMonth, setHolidayMonth] = useState(() => new Date(new Date().getFullYear(), new Date().getMonth(), 1))
  const [holidays, setHolidays] = useState<{id:string;staff_email:string;staff_name:string;start_date:string;end_date:string;note:string|null}[]>([])
  const [holidayStaffId, setHolidayStaffId] = useState('')
  const [holidayStart, setHolidayStart] = useState('')
  const [holidayEnd, setHolidayEnd] = useState('')
  const [holidayNote, setHolidayNote] = useState('')
  const [mySessions, setMySessions] = useState<MySessionDuty[]>([])
  const [mySessionsLoading, setMySessionsLoading] = useState(true)
  const [selectedMySessionsDay, setSelectedMySessionsDay] = useState('')
  const myStaffLinkKey = `acm-my-staff-link-${accountEmail.trim().toLowerCase()}`
  const [myStaffId, setMyStaffId] = useState(() => localStorage.getItem(myStaffLinkKey) ?? '')
  const fileInputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    if (!programme) return
    ensureWorkingStaffForDays(
      Array.from(new Set(programme.rows.map((row) => row.day))),
    )
  }, [programme, staff])


  useEffect(() => {
    fetch('https://api.open-meteo.com/v1/forecast?latitude=52.69&longitude=0.95&current=temperature_2m,weather_code,wind_speed_10m&timezone=Europe%2FLondon')
      .then((response) => response.ok ? response.json() : Promise.reject())
      .then((data) => setWeather({
        temperature: Math.round(data.current.temperature_2m),
        wind: Math.round(data.current.wind_speed_10m),
        code: Number(data.current.weather_code),
      }))
      .catch(() => setWeather(null))
  }, [])

  useEffect(() => {
    if (!accountEmail || !staff.length) return
    async function loadStaffAvailability() {
      const { data } = await supabase.from('staff_availability').select('staff_email,day,status')
      if (!data) return
      const staffByEmail = new Map(staff.filter((member) => member.email).map((member) => [member.email!.trim().toLowerCase(), member.id]))
      const nextWorking = { ...workingByDay }
      const nextSickness = { ...sicknessByDay }
      for (const entry of data as {staff_email:string;day:string;status:'available'|'holiday'|'sick'}[]) {
        const staffId = staffByEmail.get(entry.staff_email.toLowerCase())
        if (!staffId) continue
        const working = new Set(nextWorking[entry.day] ?? staff.map((member) => member.id))
        const sick = new Set(nextSickness[entry.day] ?? [])
        if (entry.status === 'available') { working.add(staffId); sick.delete(staffId) }
        if (entry.status === 'holiday') { working.delete(staffId); sick.delete(staffId) }
        if (entry.status === 'sick') { working.delete(staffId); sick.add(staffId) }
        nextWorking[entry.day] = Array.from(working)
        nextSickness[entry.day] = Array.from(sick)
      }
      setWorkingByDay(nextWorking)
      setSicknessByDay(nextSickness)
      localStorage.setItem(WORKING_KEY, JSON.stringify(nextWorking))
      localStorage.setItem(SICKNESS_KEY, JSON.stringify(nextSickness))
    }
    loadStaffAvailability()
    const channel = supabase.channel('manager-availability-updates').on('postgres_changes', { event: '*', schema: 'public', table: 'staff_availability' }, loadStaffAvailability).subscribe()
    return () => { supabase.removeChannel(channel) }
  }, [accountEmail, staff])

  const myStaffMember = useMemo(() => {
    const email = accountEmail.trim().toLowerCase()
    const linked = staff.find((member) => member.id === myStaffId)
    if (linked) return linked
    const byEmail = staff.find((member) => member.email?.trim().toLowerCase() === email)
    if (byEmail) return byEmail
    const wantedName = normaliseIdentity(displayName)
    return wantedName
      ? staff.find((member) => normaliseIdentity(member.name) === wantedName)
      : undefined
  }, [staff, myStaffId, accountEmail, displayName])

  function linkMyStaffProfile(staffId: string) {
    setMyStaffId(staffId)
    if (staffId) localStorage.setItem(myStaffLinkKey, staffId)
    else localStorage.removeItem(myStaffLinkKey)
  }

  async function loadMySessions() {
    if (!accountEmail) return
    setMySessionsLoading(true)
    const { data, error } = await supabase
      .from('rota_assignments')
      .select('id,programme_name,day,session,activity_name,group_numbers,duty_type,school_name,building_name,party_leader_name,staff_email,staff_name')
      .order('day')
      .order('session')

    if (error) {
      setImportMessage(`Could not load your sessions: ${error.message}`)
    } else {
      const loginEmail = accountEmail.trim().toLowerCase()
      const identityNames = new Set(
        [myStaffMember?.name, displayName]
          .map(normaliseIdentity)
          .filter(Boolean),
      )
      const duties = ((data ?? []) as (MySessionDuty & { staff_email: string; staff_name: string })[])
        .filter((duty) =>
          duty.staff_email?.trim().toLowerCase() === loginEmail ||
          identityNames.has(normaliseIdentity(duty.staff_name)),
        )
      setMySessions(duties)
      setSelectedMySessionsDay((current) =>
        current && duties.some((duty) => duty.day === current)
          ? current
          : duties[0]?.day ?? '',
      )
    }
    setMySessionsLoading(false)
  }

  useEffect(() => {
    loadMySessions()
    const channel = supabase
      .channel(`manager-my-sessions-${accountEmail}`)
      .on(
        'postgres_changes',
        { event: '*', schema: 'public', table: 'rota_assignments' },
        loadMySessions,
      )
      .subscribe()
    return () => { supabase.removeChannel(channel) }
  }, [accountEmail, myStaffMember?.id, displayName])

  async function loadHolidays() {
    const { data, error } = await supabase
      .from('staff_holidays')
      .select('id,staff_email,staff_name,start_date,end_date,note')
      .order('start_date')
    if (error) setImportMessage(error.message)
    else setHolidays((data ?? []) as typeof holidays)
  }

  useEffect(() => {
    loadHolidays()
    const channel = supabase.channel('holiday-calendar-updates')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'staff_holidays' }, loadHolidays)
      .subscribe()
    return () => { supabase.removeChannel(channel) }
  }, [])

  async function addHoliday() {
    if (!canManageHolidays) {
      setImportMessage('Team Leaders can view holidays but cannot add or edit them.')
      return
    }
    const member = staff.find((item) => item.id === holidayStaffId)
    if (!member || !holidayStart || !holidayEnd) {
      setImportMessage('Choose a staff member and holiday dates.')
      return
    }
    if (holidayEnd < holidayStart) {
      setImportMessage('The holiday end date cannot be before the start date.')
      return
    }
    const { error } = await supabase.from('staff_holidays').insert({
      staff_email: (member.email ?? '').trim().toLowerCase(),
      staff_name: member.name,
      start_date: holidayStart,
      end_date: holidayEnd,
      note: holidayNote.trim() || null,
    })
    if (error) setImportMessage(error.message)
    else {
      setHolidayStaffId(''); setHolidayStart(''); setHolidayEnd(''); setHolidayNote('')
      setImportMessage(`${member.name}'s holiday was added.`)
      loadHolidays()
    }
  }

  async function deleteHoliday(id: string) {
    if (!canManageHolidays) {
      setImportMessage('Team Leaders can view holidays but cannot delete them.')
      return
    }
    const { error } = await supabase.from('staff_holidays').delete().eq('id', id)
    if (error) setImportMessage(error.message)
    else loadHolidays()
  }

  function holidayCalendarDays() {
    const first = new Date(holidayMonth.getFullYear(), holidayMonth.getMonth(), 1)
    const start = new Date(first)
    start.setDate(first.getDate() - ((first.getDay() + 6) % 7))
    return Array.from({ length: 42 }, (_, index) => {
      const date = new Date(start); date.setDate(start.getDate() + index); return date
    })
  }

  function dateKey(date: Date) {
    const year = date.getFullYear(); const month = String(date.getMonth()+1).padStart(2,'0'); const day = String(date.getDate()).padStart(2,'0')
    return `${year}-${month}-${day}`
  }

  function sessionDemandForDay(day: string) {
    if (!programme) return []

    const sessionMap = new Map<
      string,
      { session: string; activityStaff: number; arrivalStaff: number }
    >()

    for (const row of programme.rows.filter((item) => item.day === day)) {
      const activeCells = activityCellsForRow(row)

      const current = sessionMap.get(row.session) ?? {
        session: row.session,
        activityStaff: 0,
        arrivalStaff: 0,
      }

      current.activityStaff += activeCells.length
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
        activityCellsForRow(row).map((cell) => ({ row, cell })),
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

  async function setStaffRole(staffId: string, role: StaffRole) {
    const member = staff.find((item) => item.id === staffId)
    const next = staff.map((item) =>
      item.id === staffId
        ? { ...item, role, teamLeader: role === 'teamLeader' }
        : item,
    )
    setStaff(next)
    localStorage.setItem(STAFF_KEY, JSON.stringify(next))

    if (member?.email) {
      const profileRole = role === 'centreManager'
        ? 'centreManager'
        : role === 'activityManager'
          ? 'activityManager'
          : role === 'teamLeader'
            ? 'teamLeader'
            : 'staff'
      const { error } = await supabase
        .from('profiles')
        .update({ role: profileRole })
        .eq('email', member.email.trim().toLowerCase())
      if (error) setImportMessage(`Role saved locally, but login access was not updated: ${error.message}`)
    }
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

  function qualificationIsValid(member: StaffMember, code: string) {
    return member.qualifications.includes(code)
  }

  function arrivalStaffForDaySession(day: string, session: string) {
    const used = new Set<string>()
    if (session !== '3') return used
    arrivalRows
      .filter((row) => row.day === day)
      .forEach((row) => {
        const current = arrivalAssignment(row)
        if (current.leaderId) used.add(current.leaderId)
        current.guideIds.filter(Boolean).forEach((id) => used.add(id))
      })
    return used
  }

  function activityStaffForDaySession(day: string, session: string) {
    const used = new Set<string>()
    if (!programme) return used
    programme.rows
      .filter((row) => row.day === day && row.session === session)
      .forEach((row) => activityCellsForRow(row).forEach((cell) => {
        const staffId = assignments[cellKey(row.id, cell.group)]
        if (staffId) used.add(staffId)
      }))
    return used
  }

  function aiBuildEntireRota() {
    if (!programme) return
    let next: StaffingAssignment = { ...assignments }
    const workload = new Map<string, number>()
    const sortedRows = [...programme.rows].sort((a,b) => a.day.localeCompare(b.day) || Number(a.session)-Number(b.session))
    for (const row of sortedRows) {
      const workingIds = new Set(workingByDay[row.day] ?? staff.map((m) => m.id))
      const sickIds = new Set(sicknessByDay[row.day] ?? [])
      for (const cell of activityCellsForRow(row)) {
        const key = cellKey(row.id, cell.group)
        if (next[key]) { workload.set(next[key], (workload.get(next[key]) ?? 0) + 1); continue }
        const candidates = staff.filter((m) => workingIds.has(m.id) && !sickIds.has(m.id) && qualificationIsValid(m, cell.activityCode) && !arrivalStaffForDaySession(row.day, row.session).has(m.id))
          .filter((m) => !sortedRows.some((other) => other.day === row.day && other.session === row.session && other.cells.some((c) => next[cellKey(other.id,c.group)] === m.id)))
          .sort((a,b) => (rolePriority(resolvedRole(a))-rolePriority(resolvedRole(b))) || ((workload.get(a.id)??0)-(workload.get(b.id)??0)) || a.name.localeCompare(b.name))
        if (candidates[0]) { next[key]=candidates[0].id; workload.set(candidates[0].id,(workload.get(candidates[0].id)??0)+1) }
      }
    }
    setAssignments(next)
    localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(next))
    setImportMessage('AI rota builder completed the programme using availability, valid qualifications, workload balancing and conflict prevention.')
  }

  function isWaterActivity(code: string) {
    const activity = activities.find((item) => item.code === code)
    const text = normaliseActivityText(`${code} ${activity?.name ?? ''}`)
    return ['water', 'sailing', 'kayak', 'canoe', 'paddle', 'raft', 'sup', 'windsurf'].some((term) => text.includes(term))
  }

  function hadWaterInPreviousPairedSession(staffId: string, day: string, session: string, currentAssignments: StaffingAssignment) {
    const previousSession = session === '2' ? '1' : session === '4' ? '3' : ''
    if (!previousSession || !programme) return false
    return programme.rows
      .filter((row) => row.day === day && row.session === previousSession)
      .some((row) => activityCellsForRow(row).some((cell) =>
        currentAssignments[cellKey(row.id, cell.group)] === staffId && isWaterActivity(cell.activityCode),
      ))
  }

  function autoFillStaffing(day: string) {
    if (!programme || !day) return

    const sickIds = new Set(sicknessByDay[day] ?? [])
    const workingIds = new Set(
      workingByDay[day] ?? staff.map((member) => member.id),
    )
    const nextAssignments: StaffingAssignment = { ...assignments }

    // Remove assignments for sick staff on this day before rebuilding gaps.
    programme.rows
      .filter((row) => row.day === day)
      .forEach((row) =>
        row.cells.forEach((cell) => {
          const key = cellKey(row.id, cell.group)
          const assignedId = nextAssignments[key]
          if (assignedId && sickIds.has(assignedId)) delete nextAssignments[key]
        }),
      )

    const workload = new Map<string, number>()
    Object.entries(nextAssignments).forEach(([, staffId]) => {
      workload.set(staffId, (workload.get(staffId) ?? 0) + 1)
    })

    const dayRows = programme.rows.filter((row) => row.day === day)

    for (const row of dayRows) {
      for (const cell of activityCellsForRow(row)) {

        const key = cellKey(row.id, cell.group)
        if (nextAssignments[key]) continue

        const candidates = staff
          .filter(
            (member) =>
              workingIds.has(member.id) &&
              !sickIds.has(member.id) &&
              qualificationIsValid(member, cell.activityCode) &&
              !arrivalStaffForDaySession(day, row.session).has(member.id),
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
            if (isWaterActivity(cell.activityCode) && (row.session === '2' || row.session === '4')) {
              const aWaterContinuity = hadWaterInPreviousPairedSession(a.id, day, row.session, nextAssignments)
              const bWaterContinuity = hadWaterInPreviousPairedSession(b.id, day, row.session, nextAssignments)
              if (aWaterContinuity !== bWaterContinuity) return aWaterContinuity ? -1 : 1
            }

            const priorityDifference = rolePriority(resolvedRole(a)) - rolePriority(resolvedRole(b))
            if (priorityDifference !== 0) return priorityDifference

            const workloadDifference = (workload.get(a.id) ?? 0) - (workload.get(b.id) ?? 0)
            if (workloadDifference !== 0) return workloadDifference

            return Math.random() - 0.5
          })

        const chosen = candidates[0]
        if (chosen) {
          nextAssignments[key] = chosen.id
          workload.set(chosen.id, (workload.get(chosen.id) ?? 0) + 1)
        }
      }
    }

    setAssignments(nextAssignments)
    localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(nextAssignments))
    setImportMessage(
      `Auto-filled qualified staff for ${day}. Regular staff were used first, followed by team leaders, activities manager and centre manager.`,
    )
  }

  function isDoubleBooked(
    staffId: string,
    targetRow: ProgrammeRow,
    targetGroup: number,
  ) {
    if (!programme) return false
    if (targetRow.session === '3' && arrivalStaffForDaySession(targetRow.day, targetRow.session).has(staffId)) return true
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

  function arrivalDefaults(_row: ProgrammeRow): ArrivalAssignment {
    return {
      guideIds: [],
      flatIds: [],
    }
  }

  function arrivalAssignment(row: ProgrammeRow) {
    return { ...arrivalDefaults(row), ...(arrivalAssignments[arrivalKey(row)] ?? {}) }
  }

  function updateArrivalDetails(row: ProgrammeRow, patch: Partial<ArrivalAssignment>) {
    const key = arrivalKey(row)
    const nextAssignment = { ...arrivalAssignment(row), ...patch }
    const next = { ...arrivalAssignments, [key]: nextAssignment }
    setArrivalAssignments(next)
    localStorage.setItem(ARRIVAL_ASSIGNMENTS_KEY, JSON.stringify(next))
  }

  function flatLabel(flatId: string) {
    const [building, flat] = flatId.split('-')
    return `${accommodationName(Number(building))} · Flat ${flat}`
  }

  function accommodationSummary(flatIds: string[] = []) {
    if (!flatIds.length) return ''
    const byBuilding = new Map<number, number[]>()
    flatIds.forEach((flatId) => {
      const [buildingValue, flatValue] = flatId.split('-').map(Number)
      if (!buildingValue || !flatValue) return
      const flats = byBuilding.get(buildingValue) ?? []
      flats.push(flatValue)
      byBuilding.set(buildingValue, flats)
    })
    return Array.from(byBuilding.entries())
      .sort(([a], [b]) => a - b)
      .map(([building, flats]) => `${accommodationName(building)} — Flats ${flats.sort((a, b) => a - b).join(', ')}`)
      .join('; ')
  }

  function flatsUsedByOtherSchools(row: ProgrammeRow) {
    const used = new Set<string>()
    arrivalRows.forEach((otherRow) => {
      if (otherRow.id === row.id || otherRow.day !== row.day) return
      arrivalAssignment(otherRow).flatIds?.forEach((flatId) => used.add(flatId))
    })
    return used
  }

  function toggleArrivalFlat(row: ProgrammeRow, flatId: string) {
    const current = arrivalAssignment(row)
    const usedElsewhere = flatsUsedByOtherSchools(row)
    if (!current.flatIds?.includes(flatId) && usedElsewhere.has(flatId)) {
      setImportMessage(`${flatLabel(flatId)} is already allocated to another school arriving on ${row.day}.`)
      return
    }
    const flatIds = current.flatIds?.includes(flatId)
      ? current.flatIds.filter((item) => item !== flatId)
      : [...(current.flatIds ?? []), flatId]
    updateArrivalDetails(row, { flatIds })
  }

  function arrivalStaffUsedByOtherSchools(row: ProgrammeRow) {
    const used = new Set<string>()
    arrivalRows.forEach((otherRow) => {
      if (otherRow.id === row.id || otherRow.day !== row.day) return
      const otherAssignment = arrivalAssignment(otherRow)
      if (otherAssignment.leaderId) used.add(otherAssignment.leaderId)
      otherAssignment.guideIds.filter(Boolean).forEach((id) => used.add(id))
    })
    return used
  }

  function setArrivalLeader(row: ProgrammeRow, staffId: string) {
    const usedByOtherSchools = arrivalStaffUsedByOtherSchools(row)
    if (staffId && activityStaffForDaySession(row.day, '3').has(staffId)) {
      setImportMessage('That staff member is already assigned to a Session 3 activity.')
      return
    }
    if (staffId && usedByOtherSchools.has(staffId)) {
      setImportMessage('That staff member is already assigned to another school during an overlapping arrival window.')
      return
    }
    const current = arrivalAssignment(row)
    updateArrivalDetails(row, {
      leaderId: staffId || undefined,
      guideIds: current.guideIds.map((id) => (id === staffId ? '' : id)),
    })
  }

  function setArrivalGuide(row: ProgrammeRow, slotIndex: number, staffId: string) {
    const current = arrivalAssignment(row)
    const usedByOtherSchools = arrivalStaffUsedByOtherSchools(row)

    if (staffId && activityStaffForDaySession(row.day, '3').has(staffId)) {
      setImportMessage('That staff member is already assigned to a Session 3 activity.')
      return
    }
    if (staffId && staffId === current.leaderId) {
      setImportMessage('The Party Leader cannot also be assigned an accommodation group.')
      return
    }
    if (staffId && usedByOtherSchools.has(staffId)) {
      setImportMessage('That staff member is already assigned to another school during an overlapping arrival window.')
      return
    }

    const guideIds = [...current.guideIds]
    guideIds[slotIndex] = staffId
    if (staffId && guideIds.filter((id) => id === staffId).length > 2) {
      setImportMessage('One instructor can cover a maximum of two groups, both from this school.')
      return
    }
    updateArrivalDetails(row, { guideIds })
  }

  function autoFillArrivalSchool(row: ProgrammeRow) {
    const current = arrivalAssignment(row)
    if (!current.leaderId) {
      setImportMessage(`Select the Party Leader for ${arrivalSchoolName(row) || 'this school'} before using auto-fill.`)
      return
    }

    const populatedGroups = row.cells
    const workingIds = new Set(workingByDay[row.day] ?? staff.map((member) => member.id))
    const sickIds = new Set(sicknessByDay[row.day] ?? [])
    const unavailableIds = new Set<string>([current.leaderId, ...arrivalStaffUsedByOtherSchools(row), ...activityStaffForDaySession(row.day, '3')])

    const candidates = shuffled(
      staff.filter((member) =>
        workingIds.has(member.id) &&
        !sickIds.has(member.id) &&
        !unavailableIds.has(member.id) &&
        !['centreManager', 'activityManager'].includes(resolvedRole(member)) &&
        (arrivalStaffGroup === 'all' || resolvedRole(member) === arrivalStaffGroup),
      ),
    )

    const guideIds: string[] = []
    const useOnePerGroup = candidates.length >= populatedGroups.length
    const requiredGuides = useOnePerGroup ? populatedGroups.length : Math.ceil(populatedGroups.length / 2)
    const selected = candidates.slice(0, requiredGuides)
    populatedGroups.forEach((_, index) => {
      guideIds[index] = selected[useOnePerGroup ? index : Math.floor(index / 2)]?.id ?? ''
    })

    updateArrivalDetails(row, { guideIds })
    const unfilled = guideIds.filter((id) => !id).length
    setImportMessage(unfilled ? `${arrivalSchoolName(row)}: ${unfilled} group${unfilled === 1 ? '' : 's'} still need an instructor.` : `${arrivalSchoolName(row)}: accommodation staffing auto-filled.`)
  }

  function autoFillAllArrivalSchools(day: string) {
    const rows = arrivalRows.filter((row) => row.day === day)
    const missingLeader = rows.find((row) => !arrivalAssignment(row).leaderId)
    if (missingLeader) {
      setImportMessage(`Choose a Party Leader for ${arrivalSchoolName(missingLeader)} before using Auto-fill all schools.`)
      return
    }

    const workingIds = new Set(workingByDay[day] ?? staff.map((member) => member.id))
    const sickIds = new Set(sicknessByDay[day] ?? [])
    const reserved = activityStaffForDaySession(day, '3')
    rows.forEach((row) => {
      const leaderId = arrivalAssignment(row).leaderId
      if (leaderId) reserved.add(leaderId)
    })

    const candidates = shuffled(
      staff.filter((member) =>
        workingIds.has(member.id) &&
        !sickIds.has(member.id) &&
        !reserved.has(member.id) &&
        !['centreManager', 'activityManager'].includes(resolvedRole(member)) &&
        (arrivalStaffGroup === 'all' || resolvedRole(member) === arrivalStaffGroup),
      ),
    )

    let candidateIndex = 0
    const next = { ...arrivalAssignments }
    let unfilled = 0

    rows.forEach((row) => {
      const current = arrivalAssignment(row)
      const groupCount = row.cells.length
      const staffRemaining = candidates.length - candidateIndex
      const schoolsRemaining = rows.filter((item) => item.id === row.id || rows.indexOf(item) > rows.indexOf(row))
      const groupsRemainingAfter = schoolsRemaining.slice(1).reduce((sum, item) => sum + Math.ceil(item.cells.length / 2), 0)
      const canUseOnePerGroup = staffRemaining - groupsRemainingAfter >= groupCount
      const required = canUseOnePerGroup ? groupCount : Math.ceil(groupCount / 2)
      const selected = candidates.slice(candidateIndex, candidateIndex + required)
      candidateIndex += selected.length
      const guideIds = row.cells.map((_, index) => selected[canUseOnePerGroup ? index : Math.floor(index / 2)]?.id ?? '')
      unfilled += guideIds.filter((id) => !id).length
      next[arrivalKey(row)] = { ...current, guideIds }
    })

    setArrivalAssignments(next)
    localStorage.setItem(ARRIVAL_ASSIGNMENTS_KEY, JSON.stringify(next))
    setImportMessage(unfilled ? `Auto-fill completed. ${unfilled} group${unfilled === 1 ? '' : 's'} still need an instructor.` : `All ${day} school groups were auto-filled.`)
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
          qualificationIsValid(member, selectedStaffingCode) &&
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
    if (myStaffMember && !emailByStaffId.get(myStaffMember.id)) {
      emailByStaffId.set(myStaffMember.id, accountEmail.trim().toLowerCase())
    }

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
      building_name: string | null
      party_leader_name: string | null
      arrival_time: string | null
      departure_day: string | null
      departure_time: string | null
    }[] = []

    for (const row of programme.rows) {
      for (const cell of activityCellsForRow(row)) {
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
          school_name: arrivalSchoolName(row) || null,
          building_name: null,
          party_leader_name: null,
          arrival_time: null,
          departure_day: null,
          departure_time: null,
        })
      }
    }

    for (const row of arrivalRows) {
      const arrivalDetails = arrivalAssignment(row)
      const day = row.day
      const partyLeader = staff.find((item) => item.id === arrivalDetails.leaderId)
      const buildingName = accommodationSummary(arrivalDetails.flatIds)

      if (arrivalDetails.leaderId) {
        const leader = staff.find((item) => item.id === arrivalDetails.leaderId)
        const email = emailByStaffId.get(arrivalDetails.leaderId)
        if (leader && email) {
          publishedRows.push({
            programme_name: programme.title,
            day,
            session: row.session,
            activity_code: 'ARRIVAL',
            activity_name: 'Party Leader',
            group_numbers: [],
            duty_type: 'arrival_leader',
            staff_email: email,
            staff_name: leader.name,
            school_name: arrivalSchoolName(row) || null,
            building_name: buildingName || null,
            party_leader_name: leader.name,
            arrival_time: null,
            departure_day: null,
            departure_time: null,
          })
        }
      }

      const guideGroups = new Map<string, number[]>()
      arrivalDetails.guideIds.forEach((staffId, index) => {
        if (!staffId || !row.cells[index]) return
        const groups = guideGroups.get(staffId) ?? []
        groups.push(row.cells[index].group)
        guideGroups.set(staffId, groups)
      })

      guideGroups.forEach((groupNumbers, staffId) => {
        const guide = staff.find((item) => item.id === staffId)
        const email = emailByStaffId.get(staffId)
        if (!guide || !email) return

        publishedRows.push({
          programme_name: programme.title,
          day,
          session: row.session,
          activity_code: 'ARRIVAL',
          activity_name: 'Accommodation',
          group_numbers: groupNumbers,
          duty_type: 'arrival_instructor',
          staff_email: email,
          staff_name: guide.name,
          school_name: arrivalSchoolName(row) || null,
          building_name: buildingName || null,
          party_leader_name: partyLeader?.name ?? null,
          arrival_time: null,
          departure_day: null,
          departure_time: null,
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

  const arrivalRows = programme?.rows.flatMap(arrivalRowsFromProgrammeRow) ?? []

  const arrivalRowsForDay = arrivalRows.filter(
    (row) => row.day === activeStaffingDay,
  )
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

  const schoolsOnSite = new Set(programme?.rows.map(arrivalSchoolName).filter(Boolean) ?? []).size
  const availableTodayCount = activeStaffingDay ? (workingByDay[activeStaffingDay] ?? staff.map((m) => m.id)).filter((id) => !(sicknessByDay[activeStaffingDay] ?? []).includes(id)).length : staff.length
  const staffingShortages = populatedCells.filter(({row,cell}) => !assignments[cellKey(row.id,cell.group)]).length
  const selectedDayUnfilled = populatedCells.filter(
    ({ row, cell }) => row.day === activeStaffingDay && !assignments[cellKey(row.id, cell.group)],
  )
  const selectedDayCapacityShortfall = Math.max(0, (selectedDayBusiest?.total ?? 0) - availableTodayCount)

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
          <div className="brand-title-row"><h1>Adventure Centre Manager</h1><span className="release-pill">v0.30</span></div>
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
            <section className="hero command-hero">
              <div className="hero-copy">
                <div className="hero-kicker"><ShieldCheck size={17}/><span>Live operations command centre</span></div>
                <p className="eyebrow">Today at Norfolk Lakes</p>
                <h2>{programme?.title ?? 'Upload today’s programme'}</h2>
                <p>
                  Control programme changes, staffing, availability, qualifications
                  and daily operational checks from one live dashboard.
                </p>
                <div className="hero-actions">
                  <button className="hero-upload" onClick={() => fileInputRef.current?.click()}>
                    <FileSpreadsheet size={20} />
                    {programme ? 'Replace programme' : 'Upload Excel programme'}
                  </button>
                  <button className="hero-secondary" onClick={() => setPage('staffing')}>
                    <Users size={19}/> Open daily staffing
                  </button>
                </div>
              </div>
              <div className="hero-live-card">
                <span className="live-indicator"><span/>LIVE</span>
                <strong>{availableTodayCount}</strong>
                <small>staff available today</small>
                <div className={staffingShortages ? 'hero-alert warning' : 'hero-alert ready'}>
                  {staffingShortages ? `${staffingShortages} staffing gaps need attention` : 'Programme fully staffed'}
                </div>
              </div>
            </section>

            <div className="section-heading"><div><p className="eyebrow">Operational overview</p><h2>What needs your attention</h2></div><span>Live from the current programme</span></div>
            <section className="stats-grid operations-stats">
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
              <Stat icon={<Users />} value={schoolsOnSite} label="Schools on site" />
              <Stat icon={<UserRoundCheck />} value={availableTodayCount} label="Staff available" />
              <Stat icon={<CircleAlert />} value={staffingShortages} label="Staffing shortages" />
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

            <section className="manager-my-sessions compact-my-sessions panel">
              <div className="compact-my-sessions-head">
                <div><p className="eyebrow">My rota</p><h2>My Sessions</h2></div>
                {mySessions.length > 0 && (
                  <select value={selectedMySessionsDay} onChange={(event) => setSelectedMySessionsDay(event.target.value)} aria-label="Choose My Sessions day">
                    {Array.from(new Set(mySessions.map((duty) => duty.day))).map((day) => <option key={day} value={day}>{day}</option>)}
                  </select>
                )}
                <select className="my-staff-link-select" value={myStaffMember?.id ?? ''} onChange={(event) => linkMyStaffProfile(event.target.value)} aria-label="Link My Sessions to staff profile">
                  <option value="">Link my staff profile</option>
                  {staff.map((member) => <option key={member.id} value={member.id}>{member.name}</option>)}
                </select>
              </div>
              {mySessionsLoading ? (
                <div className="compact-my-sessions-empty">Loading sessions…</div>
              ) : mySessions.length === 0 ? (
                <div className="compact-my-sessions-empty"><CalendarDays size={20} /><span>No sessions found. Check the linked staff profile above.</span></div>
              ) : (
                <div className="compact-my-session-list">
                  {mySessions.filter((duty) => duty.day === selectedMySessionsDay).map((duty) => (
                    <article className={`compact-my-session-row duty-${duty.duty_type}`} key={duty.id}>
                      <strong>S{duty.session}</strong>
                      <div><b>{duty.activity_name}</b><span>{[
                        duty.school_name,
                        duty.group_numbers?.length ? `G${duty.group_numbers.join(' & G')}` : null,
                        duty.building_name,
                      ].filter(Boolean).join(' · ') || 'Published duty'}</span></div>
                      <em>{duty.duty_type === 'activity' ? 'Instructor' : duty.duty_type === 'arrival_leader' ? 'Party Leader' : 'Arrival'}</em>
                    </article>
                  ))}
                </div>
              )}
            </section>

            <div className="section-heading"><div><p className="eyebrow">Smart operations</p><h2>Run the centre faster</h2></div><span>Automation and daily assurance</span></div>
            <section className="v019-command-centre">
              <div className="ai-builder-card v019-ai-card">
                <Bot size={28} />
                <div><p className="eyebrow">AI rota builder</p><h3>Build the complete rota</h3><p>Balances workload and respects availability, qualifications and session conflicts.</p></div>
                <button className="primary" onClick={aiBuildEntireRota}><WandSparkles size={17}/> Build entire rota</button>
              </div>
              <div className="weather-card compact-weather-card">
                <CloudSun size={30} />
                <div>
                  <p className="eyebrow">Norfolk Lakes weather</p>
                  <h3>{weather ? `${weather.temperature}°C` : 'Weather unavailable'}</h3>
                  <p>{weather ? `Wind ${weather.wind} km/h · Current conditions` : 'Check again shortly.'}</p>
                </div>
              </div>
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

        {page === 'arrivals' && (
          <Panel title="Arrivals" onBack={() => setPage('dashboard')}>
            {!programme ? (
              <EmptyProgramme onUpload={() => fileInputRef.current?.click()} />
            ) : (
              <>
                <section className="arrivals-module-intro">
                  <div>
                    <p className="eyebrow">Programme-driven arrivals</p>
                    <h3>Monday, Wednesday and Friday · Session 3</h3>
                    <p>School names are taken directly from the uploaded programme. Allocate each school to accommodation, choose its Party Leader and staff the groups here.</p>
                  </div>
                  <span className="release-pill">v0.30</span>
                </section>

                <div className="day-tabs" role="tablist" aria-label="Arrival day">
                  {programmeDays.filter(isArrivalDay).map((day) => (
                    <button
                      key={day}
                      className={activeStaffingDay === day ? 'active' : ''}
                      onClick={() => setSelectedStaffingDay(day)}
                    >
                      {day}
                    </button>
                  ))}
                </div>

                {!arrivalRowsForDay.length ? (
                  <section className="empty-arrivals-state">
                    <Building2 size={34} />
                    <h3>No Session 3 school arrivals on {activeStaffingDay}</h3>
                    <p>The app creates cards only when the uploaded programme contains a named school in Session 3 on Monday, Wednesday or Friday.</p>
                  </section>
                ) : (
                  <>
                    <section className="arrival-board-heading">
                      <div>
                        <p className="eyebrow">{activeStaffingDay} arrivals</p>
                        <h3>{arrivalRowsForDay.length} school{arrivalRowsForDay.length === 1 ? '' : 's'} detected from the programme</h3>
                        <p>Choose accommodation and a Party Leader for every school, then fill each school separately or fill all schools at once.</p>
                      </div>
                      <button className="primary" disabled={arrivalRowsForDay.some((row) => !arrivalAssignment(row).leaderId)} onClick={() => autoFillAllArrivalSchools(activeStaffingDay)}><WandSparkles size={18}/>Auto-fill all schools</button>
                    </section>

                    <div className="arrival-cards-grid">
                      {arrivalRowsForDay.map((row, schoolIndex) => {
                        const populatedGroups = row.cells
                        const assignment = arrivalAssignment(row)
                        const availableToday = workingByDay[activeStaffingDay] ?? staff.map((member) => member.id)
                        const sickToday = sicknessByDay[activeStaffingDay] ?? []
                        const usedElsewhere = arrivalStaffUsedByOtherSchools(row)

                        const leaderOptions = staff.filter((member) => availableToday.includes(member.id) && !sickToday.includes(member.id) && ['staff', 'teamLeader'].includes(resolvedRole(member)) && !usedElsewhere.has(member.id) && !assignment.guideIds.includes(member.id))
                        const guideOptions = staff.filter((member) => availableToday.includes(member.id) && !sickToday.includes(member.id) && member.id !== assignment.leaderId && !usedElsewhere.has(member.id) && (arrivalStaffGroup === 'all' || resolvedRole(member) === arrivalStaffGroup))

                        return (
                          <section className={`arrival-card school-tone-${(schoolIndex % 6) + 1}`} key={row.id}>
                            <div className="arrival-card-heading">
                              <div><p className="eyebrow">Programme school · Session 3</p><h3>{arrivalSchoolName(row)}</h3><p>{populatedGroups.length} group{populatedGroups.length === 1 ? '' : 's'} detected</p></div>
                              <Building2 size={30} />
                            </div>

                            <section className="flat-allocation-section">
                              <div className="flat-allocation-heading">
                                <div><strong>Accommodation allocation</strong><span>Select any combination of flats across Kingfisher, Swan, Grebe, Bittern, Mallard and Teal.</span></div>
                                <span>{assignment.flatIds?.length ?? 0} flat{assignment.flatIds?.length === 1 ? '' : 's'} selected</span>
                              </div>
                              <div className="building-flat-grid">
                                {[1,2,3,4,5,6].map((building) => (
                                  <fieldset className="building-flat-picker" key={building}>
                                    <legend>{accommodationName(building)}</legend>
                                    {[1,2,3,4,5].map((flat) => {
                                      const flatId = `${building}-${flat}`
                                      const usedByAnotherSchool = flatsUsedByOtherSchools(row).has(flatId)
                                      const checked = assignment.flatIds?.includes(flatId) ?? false
                                      return <label className={usedByAnotherSchool && !checked ? 'flat-unavailable' : ''} key={flatId}>
                                        <input type="checkbox" checked={checked} disabled={usedByAnotherSchool && !checked} onChange={() => toggleArrivalFlat(row, flatId)} />
                                        Flat {flat}
                                      </label>
                                    })}
                                  </fieldset>
                                ))}
                              </div>
                              {assignment.flatIds?.length ? <p className="accommodation-summary">{accommodationSummary(assignment.flatIds)}</p> : <p className="accommodation-summary empty">No flats allocated yet.</p>}
                            </section>

                            <label className="party-leader-field">Party Leader
                              <select value={assignment.leaderId ?? ''} onChange={(event) => setArrivalLeader(row, event.target.value)}>
                                <option value="">Select Party Leader</option>
                                {leaderOptions.map((member) => <option key={member.id} value={member.id}>{member.name} · {roleLabel(resolvedRole(member))}</option>)}
                              </select>
                            </label>

                            <div className="arrival-group-list">
                              {populatedGroups.map((group, index) => (
                                <label key={group.group}>Group {group.group}
                                  <select value={assignment.guideIds[index] ?? ''} onChange={(event) => setArrivalGuide(row, index, event.target.value)}>
                                    <option value="">Select instructor</option>
                                    {guideOptions.filter((member) => assignment.guideIds.filter((id, guideIndex) => guideIndex !== index && id === member.id).length < 2).map((member) => <option key={member.id} value={member.id}>{member.name}</option>)}
                                  </select>
                                </label>
                              ))}
                            </div>

                            <div className="arrival-actions">
                              <button className="primary" disabled={!assignment.leaderId} onClick={() => autoFillArrivalSchool(row)}><WandSparkles size={18} />Auto-fill school</button>
                              <span>{assignment.leaderId ? 'One instructor per group where possible; maximum two groups from this school.' : 'Select the Party Leader first.'}</span>
                            </div>
                          </section>
                        )
                      })}
                    </div>
                  </>
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

                {(selectedDayUnfilled.length > 0 || selectedDayCapacityShortfall > 0) && (
                  <section className="staffing-shortage-alert" role="alert">
                    <CircleAlert size={24} />
                    <div>
                      <strong>Staffing warning for {activeStaffingDay}</strong>
                      <p>
                        {selectedDayCapacityShortfall > 0
                          ? `${selectedDayCapacityShortfall} more staff ${selectedDayCapacityShortfall === 1 ? 'member is' : 'members are'} needed for the busiest session. `
                          : ''}
                        {selectedDayUnfilled.length > 0
                          ? `${selectedDayUnfilled.length} session ${selectedDayUnfilled.length === 1 ? 'group is' : 'groups are'} still not assigned.`
                          : ''}
                      </p>
                    </div>
                  </section>
                )}

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

                <section className="staffing-module-note">
                  <div>
                    <p className="eyebrow">Activity staffing only</p>
                    <h3>School arrivals are managed separately</h3>
                    <p>Monday, Wednesday and Friday Session 3 school rows are automatically sent to the Arrivals page.</p>
                  </div>
                  <button className="secondary-action" onClick={() => setPage('arrivals')}>
                    <Building2 size={18} /> Open arrivals
                  </button>
                </section>

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

        {page === 'schoolNotes' && (
          <section className="panel school-notes-page">
            <button className="back" onClick={() => setPage('dashboard')}><ChevronLeft size={18} />Dashboard</button>
            <p className="eyebrow">School information</p>
            <h2>School Notes</h2>
            <p className="page-intro">Add operational notes for every school detected in the uploaded programme.</p>

            {!arrivalRows.length ? (
              <section className="empty-arrivals-state">
                <Building2 size={34} />
                <h3>No schools detected</h3>
                <p>Upload a programme containing school arrivals to create school note cards.</p>
              </section>
            ) : (
              <div className="school-notes-grid">
                {arrivalRows.map((row) => {
                  const assignment = arrivalAssignment(row)
                  return (
                    <article className="school-note-card" key={row.id}>
                      <div className="school-note-card-head">
                        <div>
                          <p className="eyebrow">{row.day} · Session 3</p>
                          <h3>{arrivalSchoolName(row)}</h3>
                          <span>{row.cells.length} group{row.cells.length === 1 ? '' : 's'} · Groups {row.cells.map((cell) => cell.group).join(', ')}</span>
                        </div>
                      </div>

                      <label className="school-note-field">
                        Notes about this school
                        <textarea
                          rows={5}
                          value={assignment.notes ?? ''}
                          placeholder="Add dietary information, teacher requests, behaviour notes, accessibility needs or other operational information…"
                          onChange={(event) => updateArrivalDetails(row, { notes: event.target.value })}
                        />
                      </label>

                    </article>
                  )
                })}
              </div>
            )}
          </section>
        )}

        {page === 'admin' && (
          <Panel title="Admin" onBack={() => setPage('dashboard')}>
            <section className="admin-choice-grid">
              <button className="admin-choice-card" onClick={() => setPage('staff')}>
                <Users size={34} />
                <div><h3>Staff</h3><p>Manage staff accounts, roles and availability.</p></div>
              </button>
              <button className="admin-choice-card" onClick={() => setPage('holidays')}>
                <CalendarRange size={25} />
                <div><h3>Holidays</h3><p>View and manage the staff holiday calendar.</p></div>
                <ChevronRight size={20} />
              </button>
              <button className="admin-choice-card" onClick={() => setPage('signoffs')}>
                <ShieldCheck size={34} />
                <div><h3>Sign-off</h3><p>Search staff and manage activity sign-offs.</p></div>
              </button>
            </section>
          </Panel>
        )}

        {page === 'staff' && (
          <Panel title="Staff management" onBack={() => setPage('admin')}>
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

        {page === 'holidays' && (
          <Panel title="Staff holidays" onBack={() => setPage('admin')}>
            <div className="holiday-summary">
              <div><p className="eyebrow">Holiday calendar</p><h3>{holidayMonth.toLocaleDateString('en-GB', { month: 'long', year: 'numeric' })}</h3></div>
              <div className="holiday-month-actions">
                <button className="secondary-action" onClick={() => setHolidayMonth(new Date(holidayMonth.getFullYear(), holidayMonth.getMonth()-1, 1))}>Previous</button>
                <button className="secondary-action" onClick={() => setHolidayMonth(new Date())}>Today</button>
                <button className="secondary-action" onClick={() => setHolidayMonth(new Date(holidayMonth.getFullYear(), holidayMonth.getMonth()+1, 1))}>Next</button>
              </div>
            </div>
            {canManageHolidays ? (
              <section className="holiday-editor">
                <select value={holidayStaffId} onChange={(event) => setHolidayStaffId(event.target.value)}>
                  <option value="">Select staff member</option>
                  {staff.map((member) => <option key={member.id} value={member.id}>{member.name}</option>)}
                </select>
                <input type="date" value={holidayStart} onChange={(event) => setHolidayStart(event.target.value)} />
                <input type="date" value={holidayEnd} onChange={(event) => setHolidayEnd(event.target.value)} />
                <input value={holidayNote} onChange={(event) => setHolidayNote(event.target.value)} placeholder="Note (optional)" />
                <button className="primary" onClick={addHoliday}>Add holiday</button>
              </section>
            ) : (
              <div className="warning-banner"><strong>View only:</strong> Team Leaders can see staff holidays, but only the Head of Centre and Activities Manager can make changes.</div>
            )}
            <div className="holiday-weekdays">{['Mon','Tue','Wed','Thu','Fri','Sat','Sun'].map((day) => <strong key={day}>{day}</strong>)}</div>
            <div className="holiday-calendar">
              {holidayCalendarDays().map((date) => {
                const key = dateKey(date)
                const entries = holidays.filter((holiday) => holiday.start_date <= key && holiday.end_date >= key)
                const today = key === dateKey(new Date())
                return <article key={key} className={`holiday-day ${date.getMonth() !== holidayMonth.getMonth() ? 'outside' : ''} ${today ? 'today' : ''}`}>
                  <span className="holiday-date">{date.getDate()}</span>
                  {entries.map((holiday) => <div className={`holiday-entry ${canManageHolidays ? '' : 'readonly'}`} key={holiday.id} title={holiday.note ?? ''}>
                    <span>{holiday.staff_name}</span>
                    {canManageHolidays && <button onClick={() => deleteHoliday(holiday.id)} aria-label={`Delete ${holiday.staff_name} holiday`}>×</button>}
                  </div>)}
                </article>
              })}
            </div>
          </Panel>
        )}

        {page === 'signoffs' && (
          <Panel title="Staff sign-offs" onBack={() => setPage('admin')}>
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

            <div className="search-box signoff-search">
              <Search size={18} />
              <input
                value={signoffSearch}
                onChange={(event) => setSignoffSearch(event.target.value)}
                placeholder="Search staff by name"
              />
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
              {staff.filter((member) => member.name.toLowerCase().includes(signoffSearch.trim().toLowerCase())).map((member) => (
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
                  {arrivalSchoolName(row) && (
                    <small>{arrivalSchoolName(row)}</small>
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
