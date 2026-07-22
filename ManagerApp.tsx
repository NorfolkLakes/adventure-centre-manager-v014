import { ChangeEvent, Fragment, KeyboardEvent, useEffect, useMemo, useRef, useState } from 'react'
import {
  Building2,
  CalendarDays,
  CalendarRange,
  ClipboardList,
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
  Monitor,
  Bot,
  CloudSun,
  X,
} from 'lucide-react'
import * as XLSX from 'xlsx'

type ZipEntry = {
  name: string
  method: number
  flags: number
  modTime: number
  modDate: number
  versionMade: number
  versionNeeded: number
  internalAttrs: number
  externalAttrs: number
  localExtra: Uint8Array
  centralExtra: Uint8Array
  comment: Uint8Array
  compressed: Uint8Array
  uncompressedSize: number
  crc: number
}

const zipEncoder = new TextEncoder()
const zipDecoder = new TextDecoder()

function escapeHtml(value: string): string {
  return value.replace(/[&<>'"]/g, (character) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    "'": '&#39;',
    '"': '&quot;',
  }[character] ?? character))
}

function zipU16(view: DataView, offset: number, value: number) { view.setUint16(offset, value, true) }
function zipU32(view: DataView, offset: number, value: number) { view.setUint32(offset, value >>> 0, true) }

function zipCrc32(data: Uint8Array) {
  let crc = 0xffffffff
  for (const byte of data) {
    crc ^= byte
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1))
  }
  return (crc ^ 0xffffffff) >>> 0
}

async function zipInflate(data: Uint8Array, method: number) {
  if (method === 0) return data
  if (method !== 8) throw new Error(`Unsupported XLSX compression method: ${method}`)
  const stream = new Blob([data.slice().buffer]).stream().pipeThrough(new DecompressionStream('deflate-raw'))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

async function zipDeflate(data: Uint8Array, method: number) {
  if (method === 0) return data
  const stream = new Blob([data.slice().buffer]).stream().pipeThrough(new CompressionStream('deflate-raw'))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

function readZipEntries(input: Uint8Array) {
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength)
  let eocd = input.length - 22
  while (eocd >= Math.max(0, input.length - 65557) && view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1
  if (eocd < 0) throw new Error('The staffing template is not a valid XLSX file.')
  const count = view.getUint16(eocd + 10, true)
  let cursor = view.getUint32(eocd + 16, true)
  const entries: ZipEntry[] = []
  for (let index = 0; index < count; index += 1) {
    if (view.getUint32(cursor, true) !== 0x02014b50) throw new Error('The staffing template ZIP directory is invalid.')
    const versionMade = view.getUint16(cursor + 4, true)
    const versionNeeded = view.getUint16(cursor + 6, true)
    const flags = view.getUint16(cursor + 8, true)
    const method = view.getUint16(cursor + 10, true)
    const modTime = view.getUint16(cursor + 12, true)
    const modDate = view.getUint16(cursor + 14, true)
    const crc = view.getUint32(cursor + 16, true)
    const compressedSize = view.getUint32(cursor + 20, true)
    const uncompressedSize = view.getUint32(cursor + 24, true)
    const nameLength = view.getUint16(cursor + 28, true)
    const extraLength = view.getUint16(cursor + 30, true)
    const commentLength = view.getUint16(cursor + 32, true)
    const internalAttrs = view.getUint16(cursor + 36, true)
    const externalAttrs = view.getUint32(cursor + 38, true)
    const localOffset = view.getUint32(cursor + 42, true)
    const nameBytes = input.slice(cursor + 46, cursor + 46 + nameLength)
    const name = zipDecoder.decode(nameBytes)
    const centralExtra = input.slice(cursor + 46 + nameLength, cursor + 46 + nameLength + extraLength)
    const comment = input.slice(cursor + 46 + nameLength + extraLength, cursor + 46 + nameLength + extraLength + commentLength)
    if (view.getUint32(localOffset, true) !== 0x04034b50) throw new Error('The staffing template ZIP entry is invalid.')
    const localNameLength = view.getUint16(localOffset + 26, true)
    const localExtraLength = view.getUint16(localOffset + 28, true)
    const localExtra = input.slice(localOffset + 30 + localNameLength, localOffset + 30 + localNameLength + localExtraLength)
    const dataStart = localOffset + 30 + localNameLength + localExtraLength
    entries.push({ name, method, flags, modTime, modDate, versionMade, versionNeeded, internalAttrs, externalAttrs, localExtra, centralExtra, comment, compressed: input.slice(dataStart, dataStart + compressedSize), uncompressedSize, crc })
    cursor += 46 + nameLength + extraLength + commentLength
  }
  return entries
}

async function rewriteXlsxPackage(input: Uint8Array, replacements: Record<string, Uint8Array>, excludedPaths: Set<string> = new Set()) {
  const entries = readZipEntries(input).filter((entry) => !excludedPaths.has(entry.name))
  const localParts: Uint8Array[] = []
  const centralParts: Uint8Array[] = []
  let localOffset = 0
  for (const entry of entries) {
    const replacement = replacements[entry.name]
    const raw = replacement ?? await zipInflate(entry.compressed, entry.method)
    const compressed = replacement ? await zipDeflate(raw, entry.method) : entry.compressed
    const crc = replacement ? zipCrc32(raw) : entry.crc
    const nameBytes = zipEncoder.encode(entry.name)
    const flags = entry.flags & ~0x0008
    const local = new Uint8Array(30 + nameBytes.length + entry.localExtra.length + compressed.length)
    const localView = new DataView(local.buffer)
    zipU32(localView, 0, 0x04034b50); zipU16(localView, 4, entry.versionNeeded); zipU16(localView, 6, flags); zipU16(localView, 8, entry.method)
    zipU16(localView, 10, entry.modTime); zipU16(localView, 12, entry.modDate); zipU32(localView, 14, crc); zipU32(localView, 18, compressed.length); zipU32(localView, 22, raw.length)
    zipU16(localView, 26, nameBytes.length); zipU16(localView, 28, entry.localExtra.length)
    local.set(nameBytes, 30); local.set(entry.localExtra, 30 + nameBytes.length); local.set(compressed, 30 + nameBytes.length + entry.localExtra.length)
    localParts.push(local)

    const central = new Uint8Array(46 + nameBytes.length + entry.centralExtra.length + entry.comment.length)
    const centralView = new DataView(central.buffer)
    zipU32(centralView, 0, 0x02014b50); zipU16(centralView, 4, entry.versionMade); zipU16(centralView, 6, entry.versionNeeded); zipU16(centralView, 8, flags); zipU16(centralView, 10, entry.method)
    zipU16(centralView, 12, entry.modTime); zipU16(centralView, 14, entry.modDate); zipU32(centralView, 16, crc); zipU32(centralView, 20, compressed.length); zipU32(centralView, 24, raw.length)
    zipU16(centralView, 28, nameBytes.length); zipU16(centralView, 30, entry.centralExtra.length); zipU16(centralView, 32, entry.comment.length); zipU16(centralView, 34, 0)
    zipU16(centralView, 36, entry.internalAttrs); zipU32(centralView, 38, entry.externalAttrs); zipU32(centralView, 42, localOffset)
    central.set(nameBytes, 46); central.set(entry.centralExtra, 46 + nameBytes.length); central.set(entry.comment, 46 + nameBytes.length + entry.centralExtra.length)
    centralParts.push(central)
    localOffset += local.length
  }
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0)
  const output = new Uint8Array(localOffset + centralSize + 22)
  let offset = 0
  localParts.forEach((part) => { output.set(part, offset); offset += part.length })
  centralParts.forEach((part) => { output.set(part, offset); offset += part.length })
  const end = new DataView(output.buffer)
  zipU32(end, offset, 0x06054b50); zipU16(end, offset + 4, 0); zipU16(end, offset + 6, 0); zipU16(end, offset + 8, entries.length); zipU16(end, offset + 10, entries.length)
  zipU32(end, offset + 12, centralSize); zipU32(end, offset + 16, localOffset); zipU16(end, offset + 20, 0)
  return output
}

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
  ArchivedStaff,
} from './types'

const PROGRAMME_KEY = 'acm-programme-current'
const HISTORY_KEY = 'acm-programme-history'
const STAFF_KEY = 'acm-staff'
const ASSIGNMENT_KEY = 'acm-assignments'
const WATER_SUPPORT_KEY = 'acm-water-support'
const SICKNESS_KEY = 'acm-sickness-by-day'
const WORKING_KEY = 'acm-working-by-day'
const ACTIVITIES_KEY = 'acm-activities'
const ARRIVAL_ASSIGNMENTS_KEY = 'acm-arrival-assignments'
const FORMER_STAFF_KEY = 'acm-former-staff'
const LOAN_HISTORY_KEY = 'acm-loan-history'
const STAFF_TIMELINE_KEY = 'acm-staff-timeline'
const PAYROLL_SYNC_KEY = 'acm-payroll-sync'
const STAFFING_ARCHIVES_KEY = 'acm-staffing-archives'
const PROGRAMME_BUILDER_KEY = 'acm-programme-builder-draft'
const PROGRAMME_LIBRARY_KEY = 'acm-programme-library'
const PROGRAMME_ARCHIVE_KEY = 'acm-programme-archive-v1'

type ProgrammePurchaseType = 'bargain' | 'super' | 'outdoor'
type BuilderSchool = { id: string; name: string; programmeName: string; purchaseType: ProgrammePurchaseType; arrivalDate: string; departureDate: string; notes: string; groups: number; requestedActivities: string[]; backupOption1: string; backupOption2: string; locked: boolean }
type ProgrammeBuilderDraft = {
  name: string
  startDate: string
  endDate: string
  purchaseType: ProgrammePurchaseType
  bargainSessionLimit: number
  bargainAllowedActivities: string[]
  schools: BuilderSchool[]
  assignments: Record<string, string>
  manualLocks: Record<string, boolean>
  notes: string
}

type SavedProgramme = { id: string; title: string; startDate: string; endDate: string; updatedAt: string; draft: ProgrammeBuilderDraft }
type ArchivedProgramme = ProgrammeImport & { archiveId: string; archivedAt: string }

const DEFAULT_BARGAIN_CODES = ['CANOE', 'KAYAK', 'ARCH', 'BT', 'VB', 'MO', 'OC', 'LR', 'AIR', 'CF', 'DISCO']
const ACTIVITY_CAPACITY: Record<string, number> = { CLIMB: 2, HR: 2, BT: 2, SAIL: 3, 'SAIL PB': 1, CANOE: 3, GCAN: 3, KAYAK: 3, SUP: 3, GSUP: 3, RAFT: 2, ARCH: 2, RIFLES: 2, CF: 30, DISCO: 30 }
const BUILDER_SESSIONS = ['1', '2', '3', '4', '5']
const BUILDER_DAY_NAMES = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN']
const WEEKDAY_ORDER: Record<string, number> = { MON: 1, MONDAY: 1, TUE: 2, TUES: 2, TUESDAY: 2, WED: 3, WEDNESDAY: 3, THU: 4, THUR: 4, THURS: 4, THURSDAY: 4, FRI: 5, FRIDAY: 5, SAT: 6, SATURDAY: 6, SUN: 7, SUNDAY: 7 }
function weekdayRank(value: string) { return WEEKDAY_ORDER[value.trim().toUpperCase().replace(/[^A-Z]/g, '')] ?? 99 }

function blankProgrammeBuilderDraft(): ProgrammeBuilderDraft {
  const today = new Date()
  const monday = new Date(today)
  const offset = (today.getDay() + 6) % 7
  monday.setDate(today.getDate() - offset)
  const friday = new Date(monday)
  friday.setDate(monday.getDate() + 4)
  const iso = (date: Date) => date.toISOString().slice(0, 10)
  return {
    name: '', startDate: iso(monday), endDate: iso(friday), purchaseType: 'normal',
    bargainSessionLimit: 15, bargainAllowedActivities: DEFAULT_BARGAIN_CODES,
    schools: [{ id: `school-${Date.now()}`, name: '', programmeName: '', purchaseType: 'normal', arrivalDate: iso(monday), departureDate: iso(friday), notes: '', groups: 1, requestedActivities: [], backupOption1: '', backupOption2: '', locked: false }], assignments: {}, manualLocks: {}, notes: '',
  }
}

function builderDateRange(startDate: string, endDate: string) {
  if (!startDate || !endDate) return [] as { date: string; day: string; label: string }[]
  const start = new Date(`${startDate}T12:00:00`)
  const end = new Date(`${endDate}T12:00:00`)
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end < start) return []
  const result: { date: string; day: string; label: string }[] = []
  const cursor = new Date(start)
  while (cursor <= end && result.length < 7) {
    const day = BUILDER_DAY_NAMES[(cursor.getDay() + 6) % 7]
    result.push({ date: cursor.toISOString().slice(0, 10), day, label: cursor.toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' }) })
    cursor.setDate(cursor.getDate() + 1)
  }
  return result
}

function friendlyProgrammeDateRange(startDate: string, endDate: string) {
  if (!startDate || !endDate) return 'Dates not set'
  const start = new Date(`${startDate}T12:00:00`)
  const end = new Date(`${endDate}T12:00:00`)
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return 'Dates not set'
  return `${start.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })} – ${end.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}`
}


type MySessionDuty = {
  id: string
  programme_name: string
  day: string
  session: string
  activity_name: string
  activity_code?: string
  group_numbers: number[]
  duty_type: string
  school_name: string | null
  building_name: string | null
  party_leader_name: string | null
}


type DayOffStatus = 'off' | 'hol' | 'sick' | 'am_off' | 'pm_off'
type StaffDayOff = {
  id: string
  staff_id: string
  staff_email: string
  staff_name: string
  day: string
  status: DayOffStatus
  note: string | null
}


type StaffingArchive = {
  id: string
  weekKey: string
  title: string
  sourceFileName: string
  archivedAt: string
  archivedBy: string
  programme: ProgrammeImport
  assignments: StaffingAssignment
  workingByDay: Record<string, string[]>
  sicknessByDay: Record<string, string[]>
  daysOff: StaffDayOff[]
  staff: StaffMember[]
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

function availabilityKeyForMember(member: StaffMember) {
  const email = (member.email ?? '').trim().toLowerCase()
  return email || `staff-id:${member.id}`
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
  const worksheet = workbook.Sheets[sheetName]
  const groupHeaderRow = rows[dayRow + 1] ?? []

  // Excel stores a merged heading only in the first cell of the merged range.
  // Build a lookup that repeats that displayed value across the exact merged
  // columns so school group counts follow the visible programme layout.
  const mergedValues = new Map<string, string>()
  for (const merge of worksheet['!merges'] ?? []) {
    const anchor = worksheet[XLSX.utils.encode_cell(merge.s)]
    const value = normaliseText(anchor?.w ?? anchor?.v)
    if (!value) continue

    for (let row = merge.s.r; row <= merge.e.r; row += 1) {
      for (let column = merge.s.c; column <= merge.e.c; column += 1) {
        mergedValues.set(`${row}:${column}`, value)
      }
    }
  }

  const groupColumns: { column: number; group: number }[] = []
  for (let column = 2; column < groupHeaderRow.length; column += 1) {
    const header = normaliseText(groupHeaderRow[column])
    const match = header.match(/^G?\s*(\d{1,2})$/i)
    const parsed = match ? Number(match[1]) : Number.NaN
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

    const cells = groupColumns.map(({ column, group }) => {
      const directValue = normaliseText(row[column])
      const mergedValue = mergedValues.get(`${rowIndex}:${column}`) ?? ''

      return {
        group,
        activityCode: (directValue || mergedValue).toUpperCase(),
      }
    })

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
  if (row.session !== '3' || (!isArrivalDay(row.day) && !row.schoolLabel?.trim())) {
    return []
  }

  const sortedCells = [...row.cells]
    .filter((cell) => cell.group >= 1 && cell.group <= 30)
    .sort((a, b) => a.group - b.group)

  const segments: ArrivalSchoolSegment[] = []
  let current: ArrivalSchoolSegment | null = null
  let currentUsesMergedSpan = false

  for (let index = 0; index < sortedCells.length; index += 1) {
    const cell = sortedCells[index]
    const value = cell.activityCode.trim()

    if (looksLikeSchoolName(value)) {
      const schoolName = value.replace(/\s+/g, ' ').trim()
      const schoolKey = normalisedProgrammeValue(schoolName)
      const previousKey = current
        ? normalisedProgrammeValue(current.schoolName)
        : ''
      const nextValue = sortedCells[index + 1]?.activityCode.trim() ?? ''
      const repeatedIntoNextCell =
        looksLikeSchoolName(nextValue) &&
        normalisedProgrammeValue(nextValue) === schoolKey

      if (current && previousKey === schoolKey) {
        current.cells.push({ ...cell, activityCode: schoolName })
        currentUsesMergedSpan = true
      } else {
        current = {
          schoolName,
          cells: [{ ...cell, activityCode: schoolName }],
        }
        segments.push(current)
        currentUsesMergedSpan = repeatedIntoNextCell
      }
      continue
    }

    if (value && value.toUpperCase() !== 'Z' && isKnownProgrammeActivity(value)) {
      current = null
      currentUsesMergedSpan = false
      continue
    }

    if (current && (!value || value.toUpperCase() === 'Z')) {
      // Legacy programmes sometimes put the school name in one cell followed by
      // blanks. Keep supporting that. When the workbook supplied an explicit
      // merged span, however, stop at the edge of that span instead of counting
      // every remaining blank group column as part of the school.
      if (currentUsesMergedSpan) {
        current = null
        currentUsesMergedSpan = false
      } else {
        current.cells.push({ ...cell, activityCode: current.schoolName })
      }
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
  accountRole?: 'centreManager' | 'activityManager' | 'teamLeader' | 'admin'
}) {
  const isAdminAccount = accountRole === 'admin'
  const canManageHolidays = accountRole === 'centreManager' || accountRole === 'activityManager'
  const canManageStaff = accountRole === 'centreManager' || accountRole === 'activityManager'
  const canRecordSickness = true
  const canViewLogs = canManageStaff
  const [page, setPage] = useState<Page>(() => isAdminAccount ? 'admin' : 'dashboard')
  const [adminAccountEmail, setAdminAccountEmail] = useState('')
  const [adminAccountBusy, setAdminAccountBusy] = useState(false)
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
  const [activities, setActivities] = useState<Activity[]>(() => {
    const saved = readJson<Activity[]>(ACTIVITIES_KEY, startingActivities)
    return saved.map((activity) => ({
      ...activity,
      colour: activity.colour ?? '#dce8f5',
      equipmentQuantity: activity.equipmentQuantity ?? 0,
      capacity: Math.max(1, activity.capacity ?? ACTIVITY_CAPACITY[activity.code] ?? 1),
      enabled: activity.enabled ?? true,
      notes: activity.notes ?? '',
      staffingRuleType: activity.staffingRuleType ?? ({ CF: 'per_x_groups', DISCO: 'per_x_groups', MO: 'per_x_groups', BT: 'per_x_groups' } as Record<string, Activity['staffingRuleType']>)[activity.code] ?? 'per_group',
      staffingRuleValue: Math.max(1, activity.staffingRuleValue ?? ({ CF: 4, DISCO: 5, MO: 3, BT: 2 } as Record<string, number>)[activity.code] ?? 1),
      staffingPriority: Math.min(3, Math.max(1, activity.staffingPriority ?? (['SAIL','CANOE','KAYAK','SUP','GSUP','GCAN','RAFT','CLIMB','HR','HR UP','HR GR'].includes(activity.code) ? 1 : ['CF','DISCO','MO','BT'].includes(activity.code) ? 3 : 2))),
      requiredQualifications: activity.requiredQualifications ?? '',
      leadQualification: activity.leadQualification ?? '',
      minimumInstructors: Math.max(1, activity.minimumInstructors ?? 1),
      maximumGroupSize: Math.max(0, activity.maximumGroupSize ?? 0),
      safetyBoatRequired: activity.safetyBoatRequired ?? ['SAIL','RAFT','SUP','GSUP','GCAN'].includes(activity.code),
      shareable: activity.shareable ?? ['CF','DISCO','MO','BT','QUIZ','FILM','VIDEO','TG','WG','VB'].includes(activity.code),
      weatherRestrictions: activity.weatherRestrictions ?? (['SAIL','CANOE','KAYAK','RAFT','SUP','GSUP','GCAN'].includes(activity.code) ? 'Follow water, wind, visibility, flooding and lightning procedures.' : ['HR','HR UP','HR GR','LR','CS'].includes(activity.code) ? 'Stop during lightning restrictions and follow site weather procedures.' : ''),
      setupChecklist: activity.setupChecklist ?? '',
      packAwayChecklist: activity.packAwayChecklist ?? '',
      emergencyProcedure: activity.emergencyProcedure ?? 'Stop the session, make the group safe, contact OCM/AM/HoC by radio and follow the centre emergency procedures.',
      sopReference: activity.sopReference ?? 'Norfolk Lakes SOPs – activity section',
      riskAssessmentReference: activity.riskAssessmentReference ?? 'Norfolk Lakes Activity Risk Assessments V1 2026',
    }))
  })

  const activityCapacity = (code: string) => Math.max(1, activities.find((activity) => activity.code === code)?.capacity ?? ACTIVITY_CAPACITY[code] ?? 1)
  const enabledActivities = activities.filter((activity) => activity.enabled !== false && activity.code !== 'Z')

  const [programmeBuilder, setProgrammeBuilder] = useState<ProgrammeBuilderDraft>(() => {
    const saved = readJson(PROGRAMME_BUILDER_KEY, blankProgrammeBuilderDraft())
    return {
      ...saved,
      manualLocks: saved.manualLocks ?? {},
      schools: (saved.schools ?? []).map((school) => ({
        ...school,
        programmeName: school.programmeName ?? school.name ?? '',
        purchaseType: school.purchaseType ?? saved.purchaseType ?? 'normal',
        arrivalDate: school.arrivalDate ?? saved.startDate ?? '',
        departureDate: school.departureDate ?? saved.endDate ?? '',
        notes: school.notes ?? saved.notes ?? '',
        requestedActivities: school.requestedActivities ?? [],
        backupOption1: school.backupOption1 ?? '',
        backupOption2: school.backupOption2 ?? '',
        locked: school.locked ?? false,
      })),
    }
  })
  const [programmeBuilderView, setProgrammeBuilderView] = useState<'build' | 'preview'>('build')
  const [programmeBuilderMode, setProgrammeBuilderMode] = useState<'design' | 'upload'>('design')
  const [programmeBuilderScreen, setProgrammeBuilderScreen] = useState<'library' | 'editor'>('library')
  const [programmeBuilderMessage, setProgrammeBuilderMessage] = useState('')
  const [draggedBuilderActivity, setDraggedBuilderActivity] = useState<{ key?: string; code: string; schoolId: string } | null>(null)
  const [savedProgrammes, setSavedProgrammes] = useState<SavedProgramme[]>(() => readJson(PROGRAMME_LIBRARY_KEY, []))
  const [activeSavedProgrammeId, setActiveSavedProgrammeId] = useState<string | null>(null)
  const [programmeSearch, setProgrammeSearch] = useState('')
  const [programmeMonth, setProgrammeMonth] = useState('')
  const [programmeArchive, setProgrammeArchive] = useState<ArchivedProgramme[]>(() => readJson(PROGRAMME_ARCHIVE_KEY, []))
  const [archiveSearch, setArchiveSearch] = useState('')
  const [archiveYear, setArchiveYear] = useState('')
  const [archiveMonth, setArchiveMonth] = useState('')
  const [assignments, setAssignments] = useState<StaffingAssignment>(() =>
    readJson(ASSIGNMENT_KEY, {}),
  )
  const [waterSupportAssignments, setWaterSupportAssignments] = useState<Record<string, string>>(() =>
    readJson(WATER_SUPPORT_KEY, {}),
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

  useEffect(() => {
    if (!programme) return
    const archiveId = `${programme.startDate ?? programme.importedAt.slice(0, 10)}|${programme.endDate ?? ''}|${programme.title || programme.sourceFileName}`.toLowerCase()
    setProgrammeArchive((current) => {
      const existing = current.find((item) => item.archiveId === archiveId)
      const archived: ArchivedProgramme = {
        ...programme,
        archiveId,
        archivedAt: existing?.archivedAt ?? new Date().toISOString(),
      }
      const next = [archived, ...current.filter((item) => item.archiveId !== archiveId)]
        .sort((a, b) => (b.startDate ?? b.importedAt).localeCompare(a.startDate ?? a.importedAt))
      localStorage.setItem(PROGRAMME_ARCHIVE_KEY, JSON.stringify(next))
      return next
    })
  }, [programme])

  const programmeDays = useMemo(
    () => Array.from(new Set(programme?.rows.map((row) => row.day) ?? [])),
    [programme],
  )
  const availabilityWeekDays = useMemo(() => {
    const today = new Date()
    today.setHours(12, 0, 0, 0)
    const mondayOffset = (today.getDay() + 6) % 7
    const monday = new Date(today)
    monday.setDate(today.getDate() - mondayOffset)
    return Array.from({ length: 7 }, (_, index) => {
      const date = new Date(monday)
      date.setDate(monday.getDate() + index)
      return date.toISOString().slice(0, 10)
    })
  }, [])
  const archiveYears = useMemo(() => Array.from(new Set(programmeArchive.map((item) => (item.startDate ?? item.importedAt.slice(0, 10)).slice(0, 4)))).filter(Boolean).sort((a, b) => b.localeCompare(a)), [programmeArchive])
  const filteredProgrammeArchive = useMemo(() => {
    const search = archiveSearch.trim().toLowerCase()
    return programmeArchive.filter((item) => {
      const start = item.startDate ?? item.importedAt.slice(0, 10)
      const year = start.slice(0, 4)
      const month = start.slice(5, 7)
      const schoolNames = (item.schoolDetails ?? []).map((school) => `${school.schoolName} ${school.programmeName}`).join(' ')
      const haystack = `${item.title} ${item.sourceFileName} ${schoolNames} ${start} ${item.endDate ?? ''}`.toLowerCase()
      return (!archiveYear || year === archiveYear) && (!archiveMonth || month === archiveMonth) && (!search || haystack.includes(search))
    })
  }, [programmeArchive, archiveSearch, archiveYear, archiveMonth])

  const [selectedStaffingDay, setSelectedStaffingDay] = useState('')
  const [staffingView, setStaffingView] = useState<'activity' | 'calendar' | 'availability'>('activity')
  const [staffingZoom, setStaffingZoom] = useState<number>(() => {
    const saved = Number(localStorage.getItem('acm-staffing-zoom') ?? 100)
    return Number.isFinite(saved) ? Math.min(150, Math.max(70, saved)) : 100
  })
  const staffingDayOptions = programme ? programmeDays : availabilityWeekDays
  const activeStaffingDay =
    selectedStaffingDay && staffingDayOptions.includes(selectedStaffingDay)
      ? selectedStaffingDay
      : staffingDayOptions[0] ?? ''
  const [importMessage, setImportMessage] = useState('')
  const [weather, setWeather] = useState<{ temperature: number; wind: number; code: number } | null>(null)
  const [currentTime, setCurrentTime] = useState(() => new Date())
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
  const [addingLoanStaff, setAddingLoanStaff] = useState(false)
  const [newStaffName, setNewStaffName] = useState('')
  const [newStaffRole, setNewStaffRole] = useState<StaffRole>('staff')
  const [newStaffQualifications, setNewStaffQualifications] = useState<string[]>([])
  const [newStaffStartDate, setNewStaffStartDate] = useState(() => dateKey(new Date()))
  const [newLoanEndDate, setNewLoanEndDate] = useState('')
  const [formerStaff, setFormerStaff] = useState<ArchivedStaff[]>(() => readJson(FORMER_STAFF_KEY, []))
  const [loanHistory, setLoanHistory] = useState<ArchivedStaff[]>(() => readJson(LOAN_HISTORY_KEY, []))
  const [staffTimeline, setStaffTimeline] = useState<Record<string, {date:string;event:string}[]>>(() => readJson(STAFF_TIMELINE_KEY, {}))
  const [payrollSyncAt, setPayrollSyncAt] = useState(() => localStorage.getItem(PAYROLL_SYNC_KEY) ?? '')
  const [staffingArchives, setStaffingArchives] = useState<StaffingArchive[]>(() => readJson(STAFFING_ARCHIVES_KEY, []))
  const [selectedStaffingLogMonth, setSelectedStaffingLogMonth] = useState('')
  const [holidayMonth, setHolidayMonth] = useState(() => new Date(new Date().getFullYear(), new Date().getMonth(), 1))
  const [holidays, setHolidays] = useState<{id:string;staff_email:string;staff_name:string;start_date:string;end_date:string;note:string|null}[]>([])
  const [holidayStaffId, setHolidayStaffId] = useState('')
  const [holidayStart, setHolidayStart] = useState('')
  const [holidayEnd, setHolidayEnd] = useState('')
  const [holidayNote, setHolidayNote] = useState('')
  const [holidaySummaryStaffId, setHolidaySummaryStaffId] = useState('')
  const [holidaySickDate, setHolidaySickDate] = useState('')
  const [daysOff, setDaysOff] = useState<StaffDayOff[]>([])
  const [daysOffView, setDaysOffView] = useState<'month' | 'week' | 'day'>('month')
  const [daysOffWeek, setDaysOffWeek] = useState(() => {
    const now = new Date(); const monday = new Date(now); monday.setDate(now.getDate() - ((now.getDay() + 6) % 7)); return monday
  })
  const [daysOffDay, setDaysOffDay] = useState(() => dateKey(new Date()))
  const [daysOffStaffId, setDaysOffStaffId] = useState('')
  const [daysOffStatus, setDaysOffStatus] = useState<DayOffStatus>('off')
  const [daysOffStart, setDaysOffStart] = useState('')
  const [daysOffEnd, setDaysOffEnd] = useState('')
  const [selectedDaysOffCell, setSelectedDaysOffCell] = useState<string>('')
  const [showDaysOffHelp, setShowDaysOffHelp] = useState(false)
  const weeklyCellRefs = useRef(new Map<string, HTMLDivElement>())
  const dailyCellRefs = useRef(new Map<string, HTMLDivElement>())
  const [publishedStaffDuties, setPublishedStaffDuties] = useState<{staff_email:string;staff_name:string;day:string;session:string;activity_name:string;duty_type:string}[]>([])
  const [lastSharedUpdate, setLastSharedUpdate] = useState<{updated_by_name:string;updated_by_email:string;updated_at:string;section:string} | null>(null)
  const [waterLeadLogs, setWaterLeadLogs] = useState<{id:string;created_at:string;programme_day:string;session:string;discipline:string;lead_staff_name:string;lead_staff_id:string;lead_group:number|null;overseen_groups:number[];confirmed_by_name:string;confirmed_by_email:string;permission_from:string}[]>([])
  const [logSearch, setLogSearch] = useState('')
  const [logDiscipline, setLogDiscipline] = useState('all')
  const [pendingWaterConfirmation, setPendingWaterConfirmation] = useState<null | {day:string;session:string;discipline:'canoe'|'kayak';staffId:string;leadGroup:number|null;overseenGroups:number[]}>(null)
  const sharedStateReadyRef = useRef(false)
  const applyingRemoteStateRef = useRef(false)
  const sharedSaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const sharedSaveInFlightRef = useRef(false)
  const sharedSaveQueuedRef = useRef(false)
  const [mySessions, setMySessions] = useState<MySessionDuty[]>([])
  const [mySessionsLoading, setMySessionsLoading] = useState(true)
  const [selectedMySessionsDay, setSelectedMySessionsDay] = useState('')
  const myStaffLinkKey = `acm-my-staff-link-${accountEmail.trim().toLowerCase()}`
  const [myStaffId, setMyStaffId] = useState(() => localStorage.getItem(myStaffLinkKey) ?? '')
  const fileInputRef = useRef<HTMLInputElement>(null)
  const appProgrammeInputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    if (!programme) return
    ensureWorkingStaffForDays(
      Array.from(new Set(programme.rows.map((row) => row.day))),
    )
  }, [programme, staff])


  useEffect(() => {
    const timer = window.setInterval(() => setCurrentTime(new Date()), 60_000)
    return () => window.clearInterval(timer)
  }, [])

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
      const staffByAvailabilityKey = new Map(
        staff.map((member) => [availabilityKeyForMember(member), member.id]),
      )
      const nextWorking = { ...workingByDay }
      const nextSickness = { ...sicknessByDay }
      for (const entry of data as {staff_email:string;day:string;status:'available'|'holiday'|'sick'}[]) {
        const staffId = staffByAvailabilityKey.get(entry.staff_email.trim().toLowerCase())
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
    if (staffId) {
      localStorage.setItem(myStaffLinkKey, staffId)
      setStaff((current) => current.map((member) =>
        member.id === staffId && !member.email?.trim()
          ? { ...member, email: accountEmail.trim().toLowerCase() }
          : member,
      ))
    } else localStorage.removeItem(myStaffLinkKey)
  }

  async function loadMySessions() {
    if (!accountEmail) return
    setMySessionsLoading(true)
    const { data, error } = await supabase
      .from('rota_assignments')
      .select('id,programme_name,day,session,activity_code,activity_name,group_numbers,duty_type,school_name,building_name,party_leader_name,staff_email,staff_name')
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

  async function loadPublishedStaffDuties() {
    const { data, error } = await supabase
      .from('rota_assignments')
      .select('staff_email,staff_name,day,session,activity_name,duty_type')
    if (error) setImportMessage(error.message)
    else setPublishedStaffDuties((data ?? []) as typeof publishedStaffDuties)
  }

  async function loadDaysOff() {
    const { data, error } = await supabase
      .from('staff_days_off')
      .select('id,staff_id,staff_email,staff_name,day,status,note')
      .order('day')
    if (error) setImportMessage(`Days Off could not be loaded: ${error.message}`)
    else setDaysOff((data ?? []) as StaffDayOff[])
  }

  async function saveDaysOffRange() {
    if (!canManageHolidays) { setImportMessage('Only the Head of Centre and Activities Manager can add days off.'); return }
    const member = staff.find((item) => item.id === daysOffStaffId)
    if (!member || !daysOffStart || !daysOffEnd || daysOffEnd < daysOffStart) { setImportMessage('Choose a staff member and a valid date range.'); return }
    const rows = uniqueDatesInRange(daysOffStart, daysOffEnd).map((day) => ({
      staff_id: member.id, staff_email: member.email ?? '', staff_name: member.name, day, status: daysOffStatus, note: null,
    }))

    // Update the screen immediately. Staffing selectors and summary cards read
    // this same state, so there is no delay while Supabase sends the change back.
    const previousDaysOff = daysOff
    setDaysOff((current) => {
      const affected = new Set(rows.map((row) => row.day))
      return [
        ...current.filter((entry) => !(entry.staff_id === member.id && affected.has(entry.day))),
        ...rows.map((row) => ({ ...row, id: `pending-${member.id}-${row.day}` } as StaffDayOff)),
      ]
    })

    const { error } = await supabase.from('staff_days_off').upsert(rows, { onConflict: 'staff_id,day' })
    if (error) {
      setDaysOff(previousDaysOff)
      setImportMessage(`Days Off could not be saved: ${error.message}`)
    } else {
      setImportMessage(`${member.name}'s days off were updated.`)
      await loadDaysOff()
    }
  }

  async function setSingleDayOff(member: StaffMember, day: string, status: DayOffStatus | 'working', blankRed = false) {
    if (!canManageHolidays && !(accountRole === 'teamLeader' && status === 'sick')) return

    // Staffing tabs use programme labels such as FRI, while the Days Off calendar
    // stores ISO dates. Always persist and compare the ISO date when available.
    const availabilityDay = programme && !/^\d{4}-\d{2}-\d{2}$/.test(day)
      ? dateForProgrammeDay(programme, day)
      : day

    // Optimistic update: all staffing calculations, information boxes, Auto-fill
    // and manual selection are blocked the instant the calendar value changes.
    const previousDaysOff = daysOff
    setDaysOff((current) => {
      const withoutExisting = current.filter((entry) => !(entry.staff_id === member.id && entry.day === availabilityDay))
      if (status === 'working') return withoutExisting
      return [...withoutExisting, {
        id: `pending-${member.id}-${availabilityDay}`,
        staff_id: member.id,
        staff_email: member.email ?? '',
        staff_name: member.name,
        day: availabilityDay,
        status,
        note: blankRed ? 'blank-red' : null,
      }]
    })

    const currentWorking = workingByDay[day] ?? staff.map((item) => item.id)
    const nextWorking = { ...workingByDay, [day]: [...new Set([...currentWorking, member.id])] }
    const currentSickness = sicknessByDay[day] ?? []
    const nextSickness = {
      ...sicknessByDay,
      [day]: status === 'sick'
        ? [...new Set([...currentSickness, member.id])]
        : currentSickness.filter((id) => id !== member.id),
    }
    setWorkingByDay(nextWorking)
    setSicknessByDay(nextSickness)
    localStorage.setItem(WORKING_KEY, JSON.stringify(nextWorking))
    localStorage.setItem(SICKNESS_KEY, JSON.stringify(nextSickness))

    const result = status === 'working'
      ? await supabase.from('staff_days_off').delete().eq('staff_id', member.id).eq('day', availabilityDay)
      : await supabase.from('staff_days_off').upsert({ staff_id: member.id, staff_email: member.email ?? '', staff_name: member.name, day: availabilityDay, status, note: blankRed ? 'blank-red' : null }, { onConflict: 'staff_id,day' })
    if (result.error) {
      setDaysOff(previousDaysOff)
      setImportMessage(`Availability could not be updated: ${result.error.message}`)
      return
    }

    await loadDaysOff()
    setImportMessage(`${member.name} is now ${status === 'working' ? 'working' : dayOffLabel(status)} on ${availabilityDay}.`)
  }

  function sortedDaysOffStaff() {
    const rank: Record<StaffRole, number> = { centreManager: 0, activityManager: 1, teamLeader: 2, staff: 3 }
    return [...staff].sort((a, b) => rank[resolvedRole(a)] - rank[resolvedRole(b)] || a.name.localeCompare(b.name))
  }

  function daysOffDisplayLabel(entry?: StaffDayOff) {
    if (!entry || entry.note === 'blank-red') return ''
    return dayOffLabel(entry.status)
  }

  function focusWeeklyCell(row: number, column: number) {
    const members = sortedDaysOffStaff()
    const dates = daysOffWeekDates()
    const nextRow = Math.max(0, Math.min(members.length - 1, row))
    const nextColumn = Math.max(0, Math.min(dates.length - 1, column))
    const key = `${members[nextRow]?.id}-${dateKey(dates[nextColumn])}`
    setSelectedDaysOffCell(key)
    requestAnimationFrame(() => weeklyCellRefs.current.get(key)?.focus())
  }

  function handleWeeklyCellKeyDown(event: KeyboardEvent<HTMLDivElement>, member: StaffMember, day: string, row: number, column: number) {
    const target = event.target as HTMLElement
    if (['INPUT', 'SELECT', 'TEXTAREA'].includes(target.tagName)) return
    const key = event.key.toLowerCase()
    if (event.key === 'ArrowLeft') { event.preventDefault(); focusWeeklyCell(row, column - 1); return }
    if (event.key === 'ArrowRight') { event.preventDefault(); focusWeeklyCell(row, column + 1); return }
    if (event.key === 'ArrowUp') { event.preventDefault(); focusWeeklyCell(row - 1, column); return }
    if (event.key === 'ArrowDown' || event.key === 'Enter') { event.preventDefault(); focusWeeklyCell(row + 1, column); return }
    if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); void setSingleDayOff(member, day, 'working'); return }
    const shortcuts: Record<string, DayOffStatus> = { o: 'off', h: 'hol', s: 'sick', a: 'am_off', p: 'pm_off' }
    if (key === 'r') { event.preventDefault(); void setSingleDayOff(member, day, 'sick', true); return }
    if (shortcuts[key]) { event.preventDefault(); void setSingleDayOff(member, day, shortcuts[key]); }
  }

  function focusDailyCell(row: number) {
    const members = sortedDaysOffStaff()
    const nextRow = Math.max(0, Math.min(members.length - 1, row))
    const key = `${members[nextRow]?.id}-${daysOffDay}`
    setSelectedDaysOffCell(key)
    requestAnimationFrame(() => dailyCellRefs.current.get(key)?.focus())
  }

  function handleDailyCellKeyDown(event: KeyboardEvent<HTMLDivElement>, member: StaffMember, row: number) {
    const target = event.target as HTMLElement
    if (['INPUT', 'SELECT', 'TEXTAREA'].includes(target.tagName)) return
    const key = event.key.toLowerCase()
    if (event.key === 'ArrowUp') { event.preventDefault(); focusDailyCell(row - 1); return }
    if (event.key === 'ArrowDown' || event.key === 'Enter') { event.preventDefault(); focusDailyCell(row + 1); return }
    if (event.key === 'ArrowLeft') { event.preventDefault(); const d = parseDateKey(daysOffDay); d.setDate(d.getDate() - 1); setDaysOffDay(dateKey(d)); return }
    if (event.key === 'ArrowRight' || event.key === 'Tab') { event.preventDefault(); const d = parseDateKey(daysOffDay); d.setDate(d.getDate() + 1); setDaysOffDay(dateKey(d)); return }
    if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); void setSingleDayOff(member, daysOffDay, 'working'); return }
    const shortcuts: Record<string, DayOffStatus> = { o: 'off', h: 'hol', s: 'sick', a: 'am_off', p: 'pm_off' }
    if (key === 'r') { event.preventDefault(); void setSingleDayOff(member, daysOffDay, 'sick', true); return }
    if (shortcuts[key]) { event.preventDefault(); void setSingleDayOff(member, daysOffDay, shortcuts[key]); }
  }

  function daysOffWeekDates() {
    return Array.from({ length: 7 }, (_, index) => { const d = new Date(daysOffWeek); d.setDate(daysOffWeek.getDate()+index); return d })
  }

  function dayOffLabel(status: DayOffStatus) {
    return status === 'am_off' ? 'AM OFF' : status === 'pm_off' ? 'PM OFF' : status.toUpperCase()
  }


  function downloadDaysOffExcel(view: 'month' | 'week' | 'day') {
    const escapeXml = (value: unknown) => String(value ?? '')
      .replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;')
    const statusStyle = (status?: DayOffStatus) => status ? `status-${status}` : 'working'
    const cell = (value: unknown, style = 'grid', mergeAcross = 0) =>
      `<Cell ss:StyleID="${style}"${mergeAcross ? ` ss:MergeAcross="${mergeAcross}"` : ''}><Data ss:Type="String">${escapeXml(value)}</Data></Cell>`
    const row = (cells: string[], height = 22) => `<Row ss:Height="${height}">${cells.join('')}</Row>`
    const workbookStart = `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">
<Styles>
<Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11"/></Style>
<Style ss:ID="title"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/></Borders><Font ss:FontName="Arial" ss:Size="16" ss:Bold="1"/></Style>
<Style ss:ID="header"><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1"/><Interior ss:Color="#E7E6E6" ss:Pattern="Solid"/></Style>
<Style ss:ID="grid"><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders><Font ss:FontName="Arial" ss:Size="11"/></Style>
<Style ss:ID="staff"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1"/></Style>
<Style ss:ID="working"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
<Style ss:ID="status-off"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders><Interior ss:Color="#FFF200" ss:Pattern="Solid"/></Style>
<Style ss:ID="status-hol"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders><Interior ss:Color="#00B050" ss:Pattern="Solid"/></Style>
<Style ss:ID="status-sick"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders><Interior ss:Color="#FF6666" ss:Pattern="Solid"/></Style>
<Style ss:ID="status-am_off"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders><Interior ss:Color="#FFF200" ss:Pattern="Solid"/></Style>
<Style ss:ID="status-pm_off"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders><Interior ss:Color="#FFF200" ss:Pattern="Solid"/></Style>
</Styles>`
    let sheetName = 'Days Off'
    let columns = ''
    let rows = ''
    let printLandscape = true
    let fileName = 'days-off.xls'

    if (view === 'week') {
      const dates = daysOffWeekDates()
      sheetName = 'Weekly Staffing'
      fileName = `weekly-staffing-${dateKey(dates[0])}.xls`
      columns = '<Column ss:Width="42"/><Column ss:Width="125"/>' + dates.map(() => '<Column ss:Width="86"/>').join('')
      rows += row([cell('WEEKLY STAFFING:', 'title', 8)], 30)
      rows += row([cell('', 'grid', 8)], 8)
      rows += row([cell('No.', 'header'), cell('Staff', 'header'), ...dates.map(d => cell(d.toLocaleDateString('en-GB',{weekday:'short'}).toUpperCase(),'header'))], 23)
      rows += row([cell('', 'header'), cell('', 'header'), ...dates.map(d => cell(d.toLocaleDateString('en-GB',{day:'numeric',month:'short'}),'header'))], 23)
      sortedDaysOffStaff().forEach((member, index) => {
        rows += row([cell(index+1,'grid'),cell(member.name,'staff'),...dates.map(d=>{const entry=daysOff.find(x=>x.staff_id===member.id&&x.day===dateKey(d));return cell(daysOffDisplayLabel(entry),statusStyle(entry?.status))})])
      })
    } else if (view === 'day') {
      const chosen = parseDateKey(daysOffDay)
      sheetName = 'Daily Days Off'
      fileName = `daily-days-off-${daysOffDay}.xls`
      printLandscape = false
      columns = '<Column ss:Width="42"/><Column ss:Width="140"/><Column ss:Width="100"/><Column ss:Width="190"/>'
      rows += row([cell('DAILY STAFFING:', 'title', 3)], 30)
      rows += row([cell(chosen.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long',year:'numeric'}),'header',3)],26)
      rows += row([cell('No.','header'),cell('Staff','header'),cell('Status','header'),cell('Availability','header')],24)
      sortedDaysOffStaff().forEach((member,index)=>{const entry=daysOff.find(x=>x.staff_id===member.id&&x.day===daysOffDay);const availability=!entry?'Available all day':entry.status==='am_off'?'Unavailable Sessions 1 & 2':entry.status==='pm_off'?'Unavailable Session 5':'Unavailable all day';rows+=row([cell(index+1),cell(member.name,'staff'),cell(daysOffDisplayLabel(entry),statusStyle(entry?.status)),cell(availability)])})
    } else {
      const year=holidayMonth.getFullYear(), month=holidayMonth.getMonth(), count=new Date(year,month+1,0).getDate()
      const dates=Array.from({length:count},(_,i)=>new Date(year,month,i+1))
      sheetName = 'Monthly Days Off'
      fileName = `monthly-days-off-${year}-${String(month+1).padStart(2,'0')}.xls`
      columns = '<Column ss:Width="38"/><Column ss:Width="110"/>' + dates.map(()=>'<Column ss:Width="34"/>').join('')
      rows += row([cell(`MONTHLY DAYS OFF — ${holidayMonth.toLocaleDateString('en-GB',{month:'long',year:'numeric'}).toUpperCase()}`,'title',count+1)],30)
      rows += row([cell('No.','header'),cell('Staff','header'),...dates.map(d=>cell(`${d.toLocaleDateString('en-GB',{weekday:'short'})} ${d.getDate()}`,'header'))],30)
      sortedDaysOffStaff().forEach((member,index)=>{rows+=row([cell(index+1),cell(member.name,'staff'),...dates.map(d=>{const entry=daysOff.find(x=>x.staff_id===member.id&&x.day===dateKey(d));return cell(daysOffDisplayLabel(entry),statusStyle(entry?.status))})])})
    }

    const xml = `${workbookStart}<Worksheet ss:Name="${escapeXml(sheetName)}"><Table>${columns}${rows}</Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><PageSetup><Layout x:Orientation="${printLandscape?'Landscape':'Portrait'}"/><PageMargins x:Bottom="0.35" x:Left="0.25" x:Right="0.25" x:Top="0.35"/></PageSetup><FitToPage/><Print><FitWidth>1</FitWidth><FitHeight>1</FitHeight><ValidPrinterInfo/></Print><Selected/></WorksheetOptions></Worksheet></Workbook>`
    const blob = new Blob([xml], { type: 'application/vnd.ms-excel;charset=utf-8' })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a'); link.href = url; link.download = fileName; link.click()
    setTimeout(() => URL.revokeObjectURL(url), 1000)
  }


  function payrollRoleForMember(member: StaffMember) {
    const key = normaliseIdentity(member.name)
    if (key === 'ash' || key === 'ashley' || key.includes('ash hampton') || key.includes('hampton ash')) return 'activityManager' as const
    if (['jess', 'jessica', 'alice', 'connor', 'joe', 'joseph'].includes(key)) return 'teamLeader' as const
    return resolvedRole(member)
  }

  function payrollStaff(month: Date) {
    const monthEnd = new Date(month.getFullYear(), month.getMonth() + 1, 0)
    return staff.filter((member) => {
      if (member.employmentType === 'loan') return false
      if (!member.startDate) return true
      return parseDateKey(member.startDate) <= monthEnd
    })
  }

  function payrollTotalsForMember(member: StaffMember, month: Date) {
    const year = month.getFullYear(), monthIndex = month.getMonth()
    const monthStart = new Date(year, monthIndex, 1), monthEnd = new Date(year, monthIndex + 1, 0)
    const employmentStart = member.startDate ? parseDateKey(member.startDate) : monthStart
    const effectiveStart = employmentStart > monthStart ? employmentStart : monthStart
    if (effectiveStart > monthEnd) return { workedDays: 0, holidayDays: 0, sickDays: 0 }

    const entries = daysOff.filter((entry) => {
      const d = parseDateKey(entry.day)
      return entry.staff_id === member.id && d >= effectiveStart && d <= monthEnd
    })
    const entryByDay = new Map(entries.map((entry) => [entry.day, entry]))
    let workedDays = 0
    let holidayDays = 0
    let sickDays = 0

    for (let day = new Date(effectiveStart); day <= monthEnd; day.setDate(day.getDate() + 1)) {
      if (day.getDay() === 0 || day.getDay() === 6) continue
      const entry = entryByDay.get(dateKey(day))
      if (entry?.status === 'hol') holidayDays += 1
      else if (entry?.status === 'sick') sickDays += 1
      else if (!entry || entry.status === 'am_off' || entry.status === 'pm_off') workedDays += 1
    }

    return { workedDays, holidayDays, sickDays }
  }

  function payrollNameMatches(member: StaffMember, payrollName: string) {
    const memberKey = normaliseIdentity(member.name)
    const payrollKey = normaliseIdentity(payrollName)
    if (!memberKey || !payrollKey) return false
    if (memberKey === payrollKey) return true
    const aliases: Record<string, string[]> = {
      ash: ['hampton ashley', 'ashley hampton'],
      ashley: ['hampton ashley', 'ashley hampton'],
      jess: ['hulse jessica', 'jessica hulse'],
      jessica: ['hulse jessica', 'jessica hulse'],
      joe: ['dickinson joseph', 'joseph dickinson'],
      joseph: ['dickinson joseph', 'joseph dickinson'],
      issy: ['ramdin isabelle', 'isabelle ramdin'],
      isabelle: ['ramdin isabelle', 'isabelle ramdin'],
      ollie: ['smith oliver', 'oliver smith'],
      oliver: ['smith oliver', 'oliver smith'],
      tom: ['green thomas', 'thomas green'],
      'tom g': ['green thomas', 'thomas green'],
      'tom w': ['woodroffe thomas', 'thomas woodroffe'],
      sam: ['ballinger sam', 'sam ballinger'],
      'sam b': ['ballinger sam', 'sam ballinger'],
      harry: ['callaghan harry', 'harry callaghan'],
      'harry c': ['callaghan harry', 'harry callaghan'],
      rafe: ['bolt rafe', 'rafe bolt'],
      aubrey: ['rion jr aubrey', 'aubrey rion jr'],
    }
    if ((aliases[memberKey] ?? []).includes(payrollKey)) return true
    const memberTokens = memberKey.split(' ').filter((token) => token.length > 1)
    const payrollTokens = payrollKey.split(' ').filter((token) => token.length > 1)
    return memberTokens.every((token) => payrollTokens.includes(token)) || payrollTokens.every((token) => memberTokens.includes(token))
  }

  function arrayBufferToBase64(buffer: ArrayBuffer) {
    const bytes = new Uint8Array(buffer)
    let binary = ''
    for (let index = 0; index < bytes.length; index += 0x8000) {
      binary += String.fromCharCode(...bytes.subarray(index, Math.min(index + 0x8000, bytes.length)))
    }
    return btoa(binary)
  }

  function base64ToArrayBuffer(value: string) {
    const binary = atob(value)
    const bytes = new Uint8Array(binary.length)
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
    return bytes.buffer
  }

  async function downloadPayroll() {
    if (!canManageHolidays) return
    try {
      setImportMessage('Preparing payroll workbook…')
      const year = holidayMonth.getFullYear()
      const monthName = holidayMonth.toLocaleDateString('en-GB', { month: 'short' })
      const workbookStorageKey = `acm-payroll-workbook-${year}`
      const generatedMonthsKey = `acm-payroll-generated-months-${year}`
      const generatedMonths = JSON.parse(localStorage.getItem(generatedMonthsKey) ?? '[]') as string[]
      if (generatedMonths.includes(monthName)) {
        throw new Error(`The ${monthName} payroll sheet already exists in this year's payroll workbook.`)
      }

      let workbook: XLSX.WorkBook
      const storedWorkbook = localStorage.getItem(workbookStorageKey)
      if (storedWorkbook) {
        workbook = XLSX.read(base64ToArrayBuffer(storedWorkbook), { type: 'array', cellStyles: true, cellDates: true })
      } else {
        const response = await fetch(`${import.meta.env.BASE_URL}instructor-payroll-template.xlsx`)
        if (!response.ok) throw new Error('Payroll template could not be loaded.')
        workbook = XLSX.read(await response.arrayBuffer(), { type: 'array', cellStyles: true, cellDates: true })
      }

      const juneName = workbook.SheetNames.find((name) => name.toLowerCase() === 'jun')
      if (!juneName || !workbook.Sheets[juneName]?.['!ref'] || workbook.Sheets[juneName]['!ref'] === 'A1') {
        throw new Error('The payroll template does not contain a populated Jun sheet.')
      }
      const juneWorksheet = workbook.Sheets[juneName]
      const targetWorksheet = structuredClone(juneWorksheet)
      workbook.Sheets[monthName] = targetWorksheet
      if (!workbook.SheetNames.includes(monthName)) workbook.SheetNames.push(monthName)

      const members = payrollStaff(holidayMonth)
      const usedMemberIds = new Set<string>()
      const roleCode: Record<ReturnType<typeof payrollRoleForMember>, string> = {
        centreManager: 'HOC', activityManager: 'AM', teamLeader: 'TL', staff: 'Inst',
      }
      const roleRows: Record<ReturnType<typeof payrollRoleForMember>, number[]> = {
        centreManager: [3], activityManager: [5], teamLeader: [8, 9, 10, 11], staff: [],
      }
      const juneRange = XLSX.utils.decode_range(juneWorksheet['!ref'] ?? 'A1:R40')
      for (let row = 13; row <= juneRange.e.r; row += 1) {
        const code = normaliseText(juneWorksheet[XLSX.utils.encode_cell({ r: row, c: 6 })]?.v)
        if (code === 'Inst') roleRows.staff.push(row + 1)
      }

      const setCell = (rowOneBased: number, column: number, value: string | number) => {
        const address = XLSX.utils.encode_cell({ r: rowOneBased - 1, c: column })
        const existing = targetWorksheet[address] ?? {}
        targetWorksheet[address] = { ...existing, t: typeof value === 'number' ? 'n' : 's', v: value, w: String(value) }
      }
      const copyRow = (sourceRowOneBased: number, targetRowOneBased: number) => {
        for (let column = 0; column < 18; column += 1) {
          const sourceAddress = XLSX.utils.encode_cell({ r: sourceRowOneBased - 1, c: column })
          const targetAddress = XLSX.utils.encode_cell({ r: targetRowOneBased - 1, c: column })
          const sourceCell = juneWorksheet[sourceAddress]
          if (sourceCell) targetWorksheet[targetAddress] = structuredClone(sourceCell)
          else delete targetWorksheet[targetAddress]
        }
        if (juneWorksheet['!rows']?.[sourceRowOneBased - 1]) {
          targetWorksheet['!rows'] ??= []
          targetWorksheet['!rows']![targetRowOneBased - 1] = structuredClone(juneWorksheet['!rows']![sourceRowOneBased - 1])
        }
      }

      const templateEntries = roleRows.centreManager.concat(roleRows.activityManager, roleRows.teamLeader, roleRows.staff)
      for (const rowOneBased of templateEntries) {
        const payrollName = normaliseText(juneWorksheet[XLSX.utils.encode_cell({ r: rowOneBased - 1, c: 1 })]?.v)
        const role = normaliseText(juneWorksheet[XLSX.utils.encode_cell({ r: rowOneBased - 1, c: 6 })]?.v)
        const matchingMember = members.find((member) => !usedMemberIds.has(member.id) && payrollNameMatches(member, payrollName))
        if (matchingMember) {
          usedMemberIds.add(matchingMember.id)
          const totals = payrollTotalsForMember(matchingMember, holidayMonth)
          setCell(rowOneBased, 8, totals.workedDays)
          setCell(rowOneBased, 9, totals.holidayDays)
          setCell(rowOneBased, 10, totals.sickDays)
        } else if (payrollName && ['HOC', 'AM', 'TL', 'Inst'].includes(role)) {
          setCell(rowOneBased, 8, 0)
          setCell(rowOneBased, 9, 0)
          setCell(rowOneBased, 10, 0)
        }
      }

      let outputLastRow = juneRange.e.r + 1
      const appendMembers = members.filter((member) => !usedMemberIds.has(member.id))
      const roleOrder = ['centreManager', 'activityManager', 'teamLeader', 'staff'] as const
      for (const role of roleOrder) {
        for (const member of appendMembers.filter((item) => payrollRoleForMember(item) === role)) {
          outputLastRow += 1
          const sourceRow = role === 'centreManager' ? 3 : role === 'activityManager' ? 5 : role === 'teamLeader' ? 8 : 14
          copyRow(sourceRow, outputLastRow)
          for (const column of [2, 3, 4, 5, 7, 11, 12, 13, 16, 17]) setCell(outputLastRow, column, '')
          setCell(outputLastRow, 0, outputLastRow - 2)
          setCell(outputLastRow, 1, member.name)
          setCell(outputLastRow, 6, roleCode[role])
          const totals = payrollTotalsForMember(member, holidayMonth)
          setCell(outputLastRow, 8, totals.workedDays)
          setCell(outputLastRow, 9, totals.holidayDays)
          setCell(outputLastRow, 10, totals.sickDays)
          usedMemberIds.add(member.id)
        }
      }

      targetWorksheet['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: outputLastRow - 1, c: 17 } })
      generatedMonths.push(monthName)
      localStorage.setItem(generatedMonthsKey, JSON.stringify(generatedMonths))

      const output = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', cellStyles: true }) as ArrayBuffer
      localStorage.setItem(workbookStorageKey, arrayBufferToBase64(output))
      const url = URL.createObjectURL(new Blob([output], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }))
      const link = document.createElement('a')
      link.href = url
      link.download = `Norfolk-Lakes-Payroll-${year}.xlsx`
      link.click()
      setTimeout(() => URL.revokeObjectURL(url), 1000)
      setImportMessage(`${monthName} was added to Norfolk-Lakes-Payroll-${year}.xlsx. The Jun names and layout were retained, and new staff were calculated from their employment start date to month end.`)
    } catch (error) {
      setImportMessage(error instanceof Error ? error.message : 'Payroll could not be downloaded.')
    }
  }


  async function loadHolidays() {
    const { data, error } = await supabase
      .from('staff_holidays')
      .select('id,staff_email,staff_name,start_date,end_date,note')
      .order('start_date')
    if (error) setImportMessage(error.message)
    else setHolidays((data ?? []) as typeof holidays)
  }

  async function loadWaterLeadLogs() {
    if (!canViewLogs) return
    const { data, error } = await supabase
      .from('water_lead_logs')
      .select('id,created_at,programme_day,session,discipline,lead_staff_name,lead_staff_id,lead_group,overseen_groups,confirmed_by_name,confirmed_by_email,permission_from')
      .order('created_at', { ascending: false })
    if (error) setImportMessage(error.message)
    else setWaterLeadLogs((data ?? []) as typeof waterLeadLogs)
  }

  useEffect(() => {
    loadHolidays()
    loadDaysOff()
    loadPublishedStaffDuties()
    loadWaterLeadLogs()
    const channel = supabase.channel('holiday-calendar-updates')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'staff_holidays' }, loadHolidays)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'staff_days_off' }, loadDaysOff)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'rota_assignments' }, loadPublishedStaffDuties)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'water_lead_logs' }, loadWaterLeadLogs)
      .subscribe()
    return () => { supabase.removeChannel(channel) }
  }, [])

  useEffect(() => {
    async function loadSharedState() {
      const { data, error } = await supabase
        .from('app_live_state')
        .select('state,updated_by_name,updated_by_email,updated_at,section')
        .eq('id', 'main')
        .maybeSingle()
      if (error) {
        setImportMessage(`Live updates could not load: ${error.message}`)
        sharedStateReadyRef.current = true
        return
      }
      if (data?.state) {
        applyingRemoteStateRef.current = true
        const state = data.state as any
        if (state.programme !== undefined) setProgramme(state.programme)
        if (state.staff) setStaff(state.staff)
        if (state.activities) setActivities(state.activities)
        if (state.assignments) setAssignments(state.assignments)
        if (state.waterSupportAssignments) setWaterSupportAssignments(state.waterSupportAssignments)
        if (state.workingByDay) setWorkingByDay(state.workingByDay)
        if (state.sicknessByDay) setSicknessByDay(state.sicknessByDay)
        if (state.arrivalAssignments) setArrivalAssignments(state.arrivalAssignments)
        if (state.staffingArchives) setStaffingArchives(state.staffingArchives)
        setLastSharedUpdate({ updated_by_name: data.updated_by_name, updated_by_email: data.updated_by_email, updated_at: data.updated_at, section: data.section })
        window.setTimeout(() => { applyingRemoteStateRef.current = false }, 0)
      }
      sharedStateReadyRef.current = true
    }
    loadSharedState()
    const channel = supabase.channel('app-live-state')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'app_live_state', filter: 'id=eq.main' }, (payload) => {
        const data = payload.new as any
        if (!data?.state || data.updated_by_email?.toLowerCase() === accountEmail.trim().toLowerCase()) return
        applyingRemoteStateRef.current = true
        const state = data.state
        if (state.programme !== undefined) setProgramme(state.programme)
        if (state.staff) setStaff(state.staff)
        if (state.activities) setActivities(state.activities)
        if (state.assignments) setAssignments(state.assignments)
        if (state.waterSupportAssignments) setWaterSupportAssignments(state.waterSupportAssignments)
        if (state.workingByDay) setWorkingByDay(state.workingByDay)
        if (state.sicknessByDay) setSicknessByDay(state.sicknessByDay)
        if (state.arrivalAssignments) setArrivalAssignments(state.arrivalAssignments)
        if (state.staffingArchives) setStaffingArchives(state.staffingArchives)
        setLastSharedUpdate({ updated_by_name: data.updated_by_name, updated_by_email: data.updated_by_email, updated_at: data.updated_at, section: data.section })
        window.setTimeout(() => { applyingRemoteStateRef.current = false }, 0)
      }).subscribe()
    return () => { supabase.removeChannel(channel) }
  }, [accountEmail])

  useEffect(() => {
    if (!sharedStateReadyRef.current || applyingRemoteStateRef.current) return
    if (sharedSaveTimerRef.current) clearTimeout(sharedSaveTimerRef.current)

    const saveSharedState = async () => {
      if (sharedSaveInFlightRef.current) {
        sharedSaveQueuedRef.current = true
        return
      }

      sharedSaveInFlightRef.current = true
      const updatedAt = new Date().toISOString()
      const updatedByName = displayName?.trim() || accountEmail
      // Archives are intentionally excluded from the live row. They can become very
      // large and were causing Supabase statement timeouts on every small rota edit.
      const state = { programme, staff, activities, assignments, waterSupportAssignments, workingByDay, sicknessByDay, arrivalAssignments }
      let lastError: { message: string } | null = null

      for (let attempt = 0; attempt < 2; attempt += 1) {
        const { error } = await supabase.from('app_live_state').upsert({
          id: 'main', state, updated_by_name: updatedByName,
          updated_by_email: accountEmail.trim().toLowerCase(), updated_at: updatedAt,
          section: page,
        }, { onConflict: 'id' })
        if (!error) {
          lastError = null
          break
        }
        lastError = error
        if (attempt === 0) await new Promise((resolve) => window.setTimeout(resolve, 1200))
      }

      sharedSaveInFlightRef.current = false
      if (lastError) {
        setImportMessage(`Live update failed: ${lastError.message}`)
      } else {
        setImportMessage((message) => message.startsWith('Live update failed:') ? '' : message)
        setLastSharedUpdate({ updated_by_name: updatedByName, updated_by_email: accountEmail, updated_at: updatedAt, section: page })
      }

      if (sharedSaveQueuedRef.current) {
        sharedSaveQueuedRef.current = false
        window.setTimeout(saveSharedState, 500)
      }
    }

    sharedSaveTimerRef.current = setTimeout(saveSharedState, 2500)
    return () => { if (sharedSaveTimerRef.current) clearTimeout(sharedSaveTimerRef.current) }
  }, [programme, staff, activities, assignments, waterSupportAssignments, workingByDay, sicknessByDay, arrivalAssignments, accountEmail, displayName, page])

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



  async function setHolidayPageSickness() {
    if (!canRecordSickness) return
    const member = staff.find((item) => item.id === holidaySummaryStaffId)
    if (!member || !holidaySickDate) {
      setImportMessage('Select a staff member and sickness date.')
      return
    }
    const availabilityKey = availabilityKeyForMember(member)
    const alreadySick = (sicknessByDay[holidaySickDate] ?? []).includes(member.id)
    const { error } = await supabase.from('staff_availability').upsert({
      staff_email: availabilityKey,
      day: holidaySickDate,
      status: alreadySick ? 'available' : 'sick',
      updated_at: new Date().toISOString(),
    }, { onConflict: 'staff_email,day' })
    if (error) {
      setImportMessage(error.message)
      return
    }
    const current = sicknessByDay[holidaySickDate] ?? []
    const nextForDay = alreadySick
      ? current.filter((id) => id !== member.id)
      : [...new Set([...current, member.id])]
    const next = { ...sicknessByDay, [holidaySickDate]: nextForDay }
    setSicknessByDay(next)
    localStorage.setItem(SICKNESS_KEY, JSON.stringify(next))
    setImportMessage(`${member.name} marked ${alreadySick ? 'available' : 'sick'} on ${holidaySickDate}.`)
  }

  function uniqueDatesInRange(start: string, end: string) {
    const dates: string[] = []
    const cursor = new Date(`${start}T12:00:00`)
    const finish = new Date(`${end}T12:00:00`)
    while (cursor <= finish) {
      dates.push(dateKey(cursor))
      cursor.setDate(cursor.getDate() + 1)
    }
    return dates
  }

  const holidaySummaryMember = staff.find((member) => member.id === holidaySummaryStaffId)
  const holidayStaffSummary = useMemo(() => {
    if (!holidaySummaryMember) return null
    const email = (holidaySummaryMember.email ?? '').trim().toLowerCase()
    const name = holidaySummaryMember.name.trim().toLowerCase()
    const duties = publishedStaffDuties.filter((duty) =>
      (email && duty.staff_email.trim().toLowerCase() === email) ||
      duty.staff_name.trim().toLowerCase() === name,
    )
    const daysWorked = new Set(duties.map((duty) => duty.day)).size
    const sessionsWorked = new Set(duties.map((duty) => `${duty.day}|${duty.session}`)).size
    const activityCounts = new Map<string, number>()
    duties
      .filter((duty) => duty.duty_type === 'activity' && duty.activity_name.trim())
      .forEach((duty) => activityCounts.set(duty.activity_name, (activityCounts.get(duty.activity_name) ?? 0) + 1))
    const ranked = [...activityCounts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    const mostRun = ranked[0] ?? null
    const leastRun = ranked.length ? [...ranked].sort((a, b) => a[1] - b[1] || a[0].localeCompare(b[0]))[0] : null
    const holidayDateSet = new Set(
      holidays
        .filter((holiday) =>
          (email && holiday.staff_email.trim().toLowerCase() === email) ||
          holiday.staff_name.trim().toLowerCase() === name,
        )
        .flatMap((holiday) => uniqueDatesInRange(holiday.start_date, holiday.end_date)),
    )
    daysOff
      .filter((entry) => memberIdForDayOff(entry) === holidaySummaryMember.id && entry.status === 'hol')
      .forEach((entry) => holidayDateSet.add(entry.day))
    const sickDateSet = new Set(
      Object.entries(sicknessByDay)
        .filter(([, ids]) => ids.includes(holidaySummaryMember.id))
        .map(([day]) => day),
    )
    daysOff
      .filter((entry) => memberIdForDayOff(entry) === holidaySummaryMember.id && entry.status === 'sick')
      .forEach((entry) => sickDateSet.add(entry.day))
    return { daysWorked, sessionsWorked, mostRun, leastRun, holidayDays: holidayDateSet.size, sickDays: sickDateSet.size }
  }, [holidaySummaryMember, publishedStaffDuties, holidays, sicknessByDay, daysOff, staff])

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

  function parseDateKey(value: string) {
    const [year, month, day] = value.split('-').map(Number)
    return new Date(year, month - 1, day)
  }

  function memberIdForDayOff(entry: StaffDayOff): string | null {
    if (staff.some((member) => member.id === entry.staff_id)) return entry.staff_id
    const email = entry.staff_email.trim().toLowerCase()
    const name = normaliseIdentity(entry.staff_name)
    return staff.find((member) =>
      (email && member.email?.trim().toLowerCase() === email) ||
      normaliseIdentity(member.name) === name,
    )?.id ?? null
  }

  function unavailableStaffIdsForSession(day: string, session?: string): Set<string> {
    const availabilityDay = programme ? dateForProgrammeDay(programme, day) : day
    // staff_days_off is the single live source of truth. Legacy local
    // sickness maps are intentionally ignored so old cached values cannot
    // keep somebody marked sick after the calendar is changed to Working.
    const ids = new Set<string>()
    holidays
      .filter((holiday) => holiday.start_date <= availabilityDay && holiday.end_date >= availabilityDay)
      .forEach((holiday) => {
        const email = holiday.staff_email.trim().toLowerCase()
        const name = normaliseIdentity(holiday.staff_name)
        const member = staff.find((item) => (email && item.email?.trim().toLowerCase() === email) || normaliseIdentity(item.name) === name)
        if (member) ids.add(member.id)
      })
    daysOff.filter((entry) => entry.day === availabilityDay).forEach((entry) => {
      const memberId = memberIdForDayOff(entry)
      if (!memberId) return
      const wholeDay = ['off','hol','sick'].includes(entry.status)
      const amBlocked = entry.status === 'am_off' && ['1','2'].includes(session ?? '')
      const pmBlocked = entry.status === 'pm_off' && session === '5'
      if (wholeDay || amBlocked || pmBlocked) ids.add(memberId)
    })
    return ids
  }

  function unavailableStaffIdsForDay(day: string): Set<string> {
    const ids = new Set<string>()
    for (const session of ['1','2','3','4','5']) {
      unavailableStaffIdsForSession(day, session).forEach((id) => ids.add(id))
    }
    return ids
  }

  function sessionDemandForDay(day: string) {
    if (!programme) return []

    const sessionMap = new Map<
      string,
      { session: string; activityStaff: number; arrivalStaff: number; waterSupportStaff: number }
    >()

    for (const cluster of staffingClustersForRows(programme.rows.filter((item) => item.day === day))) {
      const current = sessionMap.get(cluster.session) ?? {
        session: cluster.session,
        activityStaff: 0,
        arrivalStaff: 0,
        waterSupportStaff: 0,
      }
      current.activityStaff += cluster.required
      sessionMap.set(cluster.session, current)
    }

    for (const need of waterSupportNeedsForDay(day)) {
      const current = sessionMap.get(need.session) ?? { session: need.session, activityStaff: 0, arrivalStaff: 0, waterSupportStaff: 0 }
      const supportId = waterSupportAssignments[waterSupportKey(day, need.session, need.discipline)]
      const leadAlreadyRunsGroup = Boolean(supportId && disciplineAssignments(day, need.session, need.discipline).some((item) => item.staffId === supportId))
      current.waterSupportStaff += leadAlreadyRunsGroup ? 0 : 1
      sessionMap.set(need.session, current)
    }

    return Array.from(sessionMap.values())
      .map((item) => ({
        ...item,
        total: Math.max(item.activityStaff, item.arrivalStaff) + item.waterSupportStaff,
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

  function staffingRuleForActivity(code: string) {
    const activity = activities.find((item) => item.code === code)
    return {
      type: activity?.staffingRuleType ?? 'per_group',
      value: Math.max(1, activity?.staffingRuleValue ?? 1),
      priority: Math.min(3, Math.max(1, activity?.staffingPriority ?? 2)),
    }
  }

  function requiredStaffForActivity(code: string, groupCount: number) {
    const rule = staffingRuleForActivity(code)
    if (rule.type === 'fixed') return Math.max(1, rule.value)
    if (rule.type === 'per_x_groups') return Math.max(1, Math.ceil(groupCount / rule.value))
    // Manual activities remain one-per-group until a manager assigns them.
    return Math.max(1, groupCount)
  }

  type StaffingCluster = { key: string; day: string; session: string; school: string; activityCode: string; cells: { row: ProgrammeRow; group: number }[]; required: number; priority: number }

  function staffingClustersForRows(rows: ProgrammeRow[]) {
    const map = new Map<string, StaffingCluster>()
    rows.forEach((row) => activityCellsForRow(row).forEach((cell) => {
      const school = row.schoolLabel?.trim() || 'Centre programme'
      const key = `${row.day}::${row.session}::${school}::${cell.activityCode}`
      const current = map.get(key) ?? { key, day: row.day, session: row.session, school, activityCode: cell.activityCode, cells: [], required: 0, priority: staffingRuleForActivity(cell.activityCode).priority }
      current.cells.push({ row, group: cell.group })
      map.set(key, current)
    }))
    return Array.from(map.values()).map((cluster) => ({ ...cluster, required: requiredStaffForActivity(cluster.activityCode, cluster.cells.length) }))
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

  const builderDays = builderDateRange(programmeBuilder.startDate, programmeBuilder.endDate)
  const builderGroups = useMemo(() => {
    let nextGroup = 1
    return programmeBuilder.schools.flatMap((school) =>
      Array.from({ length: Math.max(1, school.groups) }, () => ({ school, group: nextGroup++ })),
    )
  }, [programmeBuilder.schools])

  function persistProgrammeLibrary(next: SavedProgramme[]) {
    setSavedProgrammes(next)
    localStorage.setItem(PROGRAMME_LIBRARY_KEY, JSON.stringify(next))
  }

  function createNewProgramme() {
    const next = blankProgrammeBuilderDraft()
    setProgrammeBuilder(next)
    localStorage.setItem(PROGRAMME_BUILDER_KEY, JSON.stringify(next))
    setActiveSavedProgrammeId(null)
    setProgrammeBuilderView('build')
    setProgrammeBuilderScreen('editor')
    setProgrammeBuilderMessage('New programme ready.')
  }

  function saveProgrammeToLibrary() {
    const title = programmeBuilder.name.trim() || programmeBuilder.schools.map((school) => school.programmeName || school.name).filter(Boolean).join(' / ') || 'Untitled programme'
    const id = activeSavedProgrammeId ?? `programme-${Date.now()}`
    const saved: SavedProgramme = { id, title, startDate: programmeBuilder.startDate, endDate: programmeBuilder.endDate, updatedAt: new Date().toISOString(), draft: programmeBuilder }
    persistProgrammeLibrary([saved, ...savedProgrammes.filter((item) => item.id !== id)])
    setActiveSavedProgrammeId(id)
    saveProgrammeBuilderDraft(programmeBuilder, `${title} saved to the programme library.`)
  }

  function openSavedProgramme(saved: SavedProgramme) {
    const draft = { ...saved.draft, manualLocks: saved.draft.manualLocks ?? {} }
    setProgrammeBuilder(draft)
    localStorage.setItem(PROGRAMME_BUILDER_KEY, JSON.stringify(draft))
    setActiveSavedProgrammeId(saved.id)
    setProgrammeBuilderView('build')
    setProgrammeBuilderScreen('editor')
    setProgrammeBuilderMessage(`${saved.title} opened. You can edit it or add another school.`)
  }

  function deleteSavedProgramme(id: string) {
    persistProgrammeLibrary(savedProgrammes.filter((item) => item.id !== id))
    if (activeSavedProgrammeId === id) setActiveSavedProgrammeId(null)
  }

  function loadSavedProgrammeIntoApp(saved: SavedProgramme) {
    openSavedProgramme(saved)
    publishProgrammeDraft(saved.draft)
  }

  const filteredSavedProgrammes = savedProgrammes.filter((saved) => {
    const search = programmeSearch.trim().toLowerCase()
    const schoolText = saved.draft.schools.map((school) => `${school.name} ${school.programmeName}`).join(' ').toLowerCase()
    const matchesSearch = !search || saved.title.toLowerCase().includes(search) || schoolText.includes(search)
    const matchesMonth = !programmeMonth || saved.startDate.startsWith(programmeMonth)
    return matchesSearch && matchesMonth
  })

  function saveProgrammeBuilderDraft(next = programmeBuilder, message = 'Draft saved.') {
    setProgrammeBuilder(next)
    localStorage.setItem(PROGRAMME_BUILDER_KEY, JSON.stringify(next))
    setProgrammeBuilderMessage(message)
  }

  function updateProgrammeBuilder(patch: Partial<ProgrammeBuilderDraft>) {
    const next = { ...programmeBuilder, ...patch }
    setProgrammeBuilder(next)
    localStorage.setItem(PROGRAMME_BUILDER_KEY, JSON.stringify(next))
    setProgrammeBuilderMessage('Draft auto-saved.')
  }

  function addBuilderSchool() {
    updateProgrammeBuilder({ schools: [...programmeBuilder.schools, { id: `school-${Date.now()}`, name: '', programmeName: '', purchaseType: 'normal', arrivalDate: programmeBuilder.startDate, departureDate: programmeBuilder.endDate, notes: '', groups: 1, requestedActivities: [], backupOption1: '', backupOption2: '', locked: false }] })
  }

  function updateBuilderSchool(id: string, patch: Partial<BuilderSchool>) {
    updateProgrammeBuilder({ schools: programmeBuilder.schools.map((school) => school.id === id ? { ...school, ...patch } : school) })
  }

  function removeBuilderSchool(id: string) {
    if (programmeBuilder.schools.length === 1) return
    updateProgrammeBuilder({ schools: programmeBuilder.schools.filter((school) => school.id !== id) })
  }

  function builderAssignmentKey(day: string, session: string, group: number) { return `${day}|${session}|${group}` }


  function normaliseBuilderDate(value: string) {
    const trimmed = value.trim()
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed
    const match = trimmed.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/)
    if (!match) return trimmed
    const [, day, month, year] = match
    return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`
  }

  function builderSchoolSessionState(school: BuilderSchool, date: string, session: string) {
    const currentDate = normaliseBuilderDate(date)
    const arrivalDate = normaliseBuilderDate(school.arrivalDate)
    const departureDate = normaliseBuilderDate(school.departureDate)

    if (!arrivalDate || !departureDate) return 'activity' as const
    if (currentDate < arrivalDate || currentDate > departureDate) return 'offsite' as const

    if (currentDate === arrivalDate) {
      if (session === '1' || session === '2') return 'offsite' as const
      if (session === '3') return 'arrival' as const
      return 'activity' as const
    }

    if (currentDate === departureDate) {
      if (session === '1' || session === '2') return 'activity' as const
      return 'departed' as const
    }

    return 'activity' as const
  }

  const WATER_ACTIVITY_CODES = new Set(['CANOE', 'KAYAK', 'RAFT', 'SUP', 'GSUP', 'GCAN', 'SAIL'])

  function builderSlotPriority(school: BuilderSchool, date: string, session: string) {
    const day = new Date(`${date}T12:00:00`).getDay()
    const isFriday = day === 5
    const isArrival = date === school.arrivalDate
    const isDeparture = date === school.departureDate
    const arrival = new Date(`${school.arrivalDate}T12:00:00`)
    const current = new Date(`${date}T12:00:00`)
    const daysAfterArrival = Math.round((current.getTime() - arrival.getTime()) / 86400000)
    if (daysAfterArrival === 1 && !isFriday && !isDeparture) return 0
    if (!isArrival && !isFriday && !isDeparture) return 1
    if (isArrival) return 2
    if (isFriday || isDeparture) return 3
    return 2
  }

  function generateSchoolAssignments(schoolId: string, preserveManual: boolean, baseAssignments = programmeBuilder.assignments) {
    const school = programmeBuilder.schools.find((item) => item.id === schoolId)
    if (!school || (school.locked && preserveManual)) return baseAssignments
    const schoolGroups = builderGroups.filter((entry) => entry.school.id === schoolId)
    const requested = school.requestedActivities.length ? school.requestedActivities : enabledActivities.map((item) => item.code)
    const allowed = school.purchaseType === 'bargain' ? requested.filter((code) => programmeBuilder.bargainAllowedActivities.includes(code)) : requested
    if (!allowed.length) return baseAssignments

    const next = { ...baseAssignments }
    const activitySlots = builderDays.flatMap((dayInfo) => BUILDER_SESSIONS
      .filter((session) => builderSchoolSessionState(school, dayInfo.date, session) === 'activity')
      .map((session) => ({ ...dayInfo, session, priority: builderSlotPriority(school, dayInfo.date, session) })))

    const activityAtCapacity = (day: string, session: string, code: string, ownKey: string) => {
      const capacity = activityCapacity(code)
      const running = Object.entries(next).filter(([key, value]) => key !== ownKey && key.startsWith(`${day}|${session}|`) && value === code).length
      return running >= capacity
    }

    // Canoeing and kayaking are one linked two-session rotation. Two groups
    // swap activities in the immediately following session. Existing manual rotations
    // are recognised as complete so Update Programme never adds an extra canoe/kayak.
    const pairedWaterGroups = new Set<number>()
    if (allowed.includes('CANOE') && allowed.includes('KAYAK')) {
      const availablePairs = [...activitySlots]
        .sort((a, b) => a.priority - b.priority || a.date.localeCompare(b.date) || Number(a.session) - Number(b.session))
        .flatMap((first) => {
          const second = activitySlots.find((candidate) => candidate.day === first.day && Number(candidate.session) === Number(first.session) + 1)
          return second ? [{ first, second }] : []
        })

      for (let index = 0; index + 1 < schoolGroups.length; index += 2) {
        const firstGroup = schoolGroups[index].group
        const secondGroup = schoolGroups[index + 1].group

        const completedManualPair = availablePairs.find(({ first, second }) => {
          const values = [
            next[builderAssignmentKey(first.day, first.session, firstGroup)],
            next[builderAssignmentKey(first.day, first.session, secondGroup)],
            next[builderAssignmentKey(second.day, second.session, firstGroup)],
            next[builderAssignmentKey(second.day, second.session, secondGroup)],
          ]
          const forward = values[0] === 'KAYAK' && values[1] === 'CANOE' && values[2] === 'CANOE' && values[3] === 'KAYAK'
          const reverse = values[0] === 'CANOE' && values[1] === 'KAYAK' && values[2] === 'KAYAK' && values[3] === 'CANOE'
          return forward || reverse
        })
        if (completedManualPair) {
          pairedWaterGroups.add(firstGroup)
          pairedWaterGroups.add(secondGroup)
          availablePairs.splice(availablePairs.indexOf(completedManualPair), 1)
          continue
        }

        const pair = availablePairs.find(({ first, second }) => {
          const keys = [
            builderAssignmentKey(first.day, first.session, firstGroup),
            builderAssignmentKey(first.day, first.session, secondGroup),
            builderAssignmentKey(second.day, second.session, firstGroup),
            builderAssignmentKey(second.day, second.session, secondGroup),
          ]
          if (preserveManual && keys.some((key) => programmeBuilder.manualLocks[key])) return false
          return !activityAtCapacity(first.day, first.session, 'KAYAK', keys[0])
            && !activityAtCapacity(first.day, first.session, 'CANOE', keys[1])
            && !activityAtCapacity(second.day, second.session, 'CANOE', keys[2])
            && !activityAtCapacity(second.day, second.session, 'KAYAK', keys[3])
        })
        if (!pair) continue
        next[builderAssignmentKey(pair.first.day, pair.first.session, firstGroup)] = 'KAYAK'
        next[builderAssignmentKey(pair.first.day, pair.first.session, secondGroup)] = 'CANOE'
        next[builderAssignmentKey(pair.second.day, pair.second.session, firstGroup)] = 'CANOE'
        next[builderAssignmentKey(pair.second.day, pair.second.session, secondGroup)] = 'KAYAK'
        pairedWaterGroups.add(firstGroup)
        pairedWaterGroups.add(secondGroup)
        availablePairs.splice(availablePairs.indexOf(pair), 1)
      }
    }

    for (const { group } of schoolGroups) {
      const unlockedSlots = activitySlots.filter((slot) => {
        const key = builderAssignmentKey(slot.day, slot.session, group)
        return !(preserveManual && programmeBuilder.manualLocks[key])
      })
      for (const slot of unlockedSlots) next[builderAssignmentKey(slot.day, slot.session, group)] = ''

      const sortedSlots = [...unlockedSlots].sort((a, b) => a.priority - b.priority || a.date.localeCompare(b.date) || Number(a.session) - Number(b.session))
      // Each selected activity may only appear once per group. Activities already
      // placed in manually locked cells (or in the linked canoe/kayak rotation)
      // are removed from the auto-fill queue.
      const alreadyAssigned = new Set(
        activitySlots
          .map((slot) => next[builderAssignmentKey(slot.day, slot.session, group)])
          .filter(Boolean),
      )
      const queue = [...allowed].filter((code) =>
        code !== 'CF' &&
        code !== 'DISCO' &&
        !alreadyAssigned.has(code) &&
        !(pairedWaterGroups.has(group) && (code === 'CANOE' || code === 'KAYAK')),
      )
      const campfireIndex = queue.indexOf('CF')
      if (campfireIndex >= 0) {
        queue.splice(campfireIndex, 1)
        const campfireSlot = [...sortedSlots].filter((slot) => slot.session === '5').sort((a, b) => b.date.localeCompare(a.date))[0]
        if (campfireSlot) {
          const key = builderAssignmentKey(campfireSlot.day, campfireSlot.session, group)
          if (!activityAtCapacity(campfireSlot.day, campfireSlot.session, 'CF', key)) {
            next[key] = 'CF'
            sortedSlots.splice(sortedSlots.indexOf(campfireSlot), 1)
          } else queue.push('CF')
        }
      }



      for (const slot of sortedSlots) {
        if (!queue.length) break
        const key = builderAssignmentKey(slot.day, slot.session, group)
        let chosenIndex = queue.findIndex((code) => code !== 'CF' && (!WATER_ACTIVITY_CODES.has(code) || slot.priority < 2) && !activityAtCapacity(slot.day, slot.session, code, key))
        if (chosenIndex < 0) chosenIndex = queue.findIndex((code) => code !== 'CF' && !activityAtCapacity(slot.day, slot.session, code, key))
        if (chosenIndex < 0) continue
        const [code] = queue.splice(chosenIndex, 1)
        next[key] = code
      }
    }
    // Campfire and Disco are whole-school evening activities. Put them together in Session 5,
    // as late in the stay as possible, without overwriting a manager's locked cell.
    const selectedCollective = allowed.includes('CF') ? 'CF' : allowed.includes('DISCO') ? 'DISCO' : ''
    if (selectedCollective) {
      const collectiveSlot = [...activitySlots].filter((slot) => slot.session === '5').sort((a, b) => b.date.localeCompare(a.date))[0]
      if (collectiveSlot) for (const { group } of schoolGroups) {
        const key = builderAssignmentKey(collectiveSlot.day, collectiveSlot.session, group)
        if (!(preserveManual && programmeBuilder.manualLocks[key])) next[key] = selectedCollective
      }
    }
    return next
  }

  function autoFillProgrammeBuilder(schoolId: string) {
    const school = programmeBuilder.schools.find((item) => item.id === schoolId)
    if (!school) return
    const next = generateSchoolAssignments(schoolId, true)
    updateProgrammeBuilder({ assignments: next })
    setProgrammeBuilderMessage(`${school.name || 'School'} updated around your manually locked sessions.`)
  }

  function updateWholeProgramme() {
    let next = { ...programmeBuilder.assignments }
    for (const school of programmeBuilder.schools) next = generateSchoolAssignments(school.id, true, next)
    updateProgrammeBuilder({ assignments: next })
    const backupSchools = programmeBuilder.schools.filter((school) => [school.backupOption1, school.backupOption2].some((code) => code && Object.values(next).includes(code)))
    setProgrammeBuilderMessage(backupSchools.length ? `Programme updated. Backup activities are in use for: ${backupSchools.map((school) => school.name || 'School').join(', ')}.` : 'Programme updated. Your manual changes were kept and the remaining sessions were rearranged around them.')
  }

  function resetProgrammeLocks() {
    updateProgrammeBuilder({ manualLocks: {} })
    setProgrammeBuilderMessage('Manual locks cleared. The next update can rearrange every activity session.')
  }

  function confirmBuilderCapacity(day: string, session: string, group: number, activityCode: string) {
    if (!activityCode) return true
    const key = builderAssignmentKey(day, session, group)
    const capacity = activityCapacity(activityCode)
    const clashes = Object.entries(programmeBuilder.assignments)
      .filter(([assignmentKey, code]) => assignmentKey !== key && assignmentKey.startsWith(`${day}|${session}|`) && code === activityCode)
      .map(([assignmentKey]) => Number(assignmentKey.split('|')[2]))
    if (clashes.length + 1 <= capacity) return true

    const affectedGroups = [...clashes, group]
    const affectedSchools = Array.from(new Set(affectedGroups.map((groupNumber) => builderGroups.find((entry) => entry.group === groupNumber)?.school.name || `Group ${groupNumber}`)))
    const activity = activities.find((item) => item.code === activityCode)
    return window.confirm(
      `${activity?.name || activityCode} capacity exceeded\n\n` +
      `${affectedSchools.join(' and ')} have ${activity?.name || activityCode} in ${day}, Session ${session}.\n` +
      `This would run ${clashes.length + 1} groups, but the maximum is ${capacity}.\n\n` +
      `Press OK to override the limit, or Cancel to keep the current programme.`
    )
  }

  function setBuilderActivity(day: string, session: string, group: number, activityCode: string) {
    if (!confirmBuilderCapacity(day, session, group, activityCode)) return
    const key = builderAssignmentKey(day, session, group)
    updateProgrammeBuilder({
      assignments: { ...programmeBuilder.assignments, [key]: activityCode },
      manualLocks: { ...programmeBuilder.manualLocks, [key]: true },
    })
    setProgrammeBuilderMessage('Manual change saved and locked. Update Programme will keep it in place.')
  }


  function dropBuilderActivity(day: string, session: string, group: number, schoolId: string) {
    if (!draggedBuilderActivity || draggedBuilderActivity.schoolId !== schoolId) return
    const targetKey = builderAssignmentKey(day, session, group)
    if (!confirmBuilderCapacity(day, session, group, draggedBuilderActivity.code)) {
      setDraggedBuilderActivity(null)
      return
    }
    const sourceKey = draggedBuilderActivity.key
    const targetValue = programmeBuilder.assignments[targetKey] ?? ''
    const nextAssignments = { ...programmeBuilder.assignments, [targetKey]: draggedBuilderActivity.code }
    const nextLocks = { ...programmeBuilder.manualLocks, [targetKey]: true }
    if (sourceKey && sourceKey !== targetKey) {
      nextAssignments[sourceKey] = targetValue
      nextLocks[sourceKey] = true
    }
    updateProgrammeBuilder({ assignments: nextAssignments, manualLocks: nextLocks })
    setDraggedBuilderActivity(null)
    setProgrammeBuilderMessage(sourceKey ? 'Activities moved. Both sessions are locked and the rest of the programme remains unchanged.' : 'Selected activity placed and locked. Press Update Programme to rebalance the remaining unlocked sessions.')
  }

  function builderValidation() {
    const issues: string[] = []
    if (!programmeBuilder.name.trim()) issues.push('Add a programme name.')
    if (!builderDays.length) issues.push('Choose a valid programme date range of up to seven days.')
    if (programmeBuilder.schools.some((school) => !school.name.trim())) issues.push('Every school needs a name.')
    if (!builderGroups.length) issues.push('Add at least one group.')
    for (const school of programmeBuilder.schools) {
      if (!school.programmeName.trim()) issues.push(`${school.name || 'Each school'} needs a programme name.`)
      if (!school.arrivalDate || !school.departureDate || school.departureDate < school.arrivalDate) issues.push(`${school.name || 'Each school'} needs valid arrival and departure dates.`)
      if (school.arrivalDate < programmeBuilder.startDate || school.departureDate > programmeBuilder.endDate) issues.push(`${school.name || 'A school'} dates must sit inside the programme week.`)
      if (school.purchaseType === 'bargain') {
        for (const { group } of builderGroups.filter((entry) => entry.school.id === school.id)) {
          const used = Object.entries(programmeBuilder.assignments).filter(([key, value]) => key.endsWith(`|${group}`) && value).map(([, value]) => value)
          if (used.some((code) => !programmeBuilder.bargainAllowedActivities.includes(code))) issues.push(`${school.name} contains an activity outside the Bargain Special package.`)
          if (used.length > programmeBuilder.bargainSessionLimit) issues.push(`Group ${group} exceeds the ${programmeBuilder.bargainSessionLimit}-session Bargain Special limit.`)
        }
      }
    }
    const slotActivities: Record<string, string[]> = {}
    Object.entries(programmeBuilder.assignments).forEach(([key, code]) => {
      if (!code) return
      const [day, session] = key.split('|')
      const slot = `${day}|${session}`
      slotActivities[slot] = [...(slotActivities[slot] ?? []), code]
      if ((code === 'CF' || code === 'DISCO') && session !== '5') issues.push(`${code === 'CF' ? 'Campfire' : 'Disco'} must be scheduled in Session 5.`)
    })
    Object.entries(slotActivities).forEach(([slot, codes]) => {
      for (const code of new Set(codes)) {
        const count = codes.filter((item) => item === code).length
        const capacity = activityCapacity(code)
        if (count > capacity) issues.push(`${code} exceeds its capacity (${count}/${capacity}) at ${slot.replace('|', ' Session ')}.`)
      }
    })
    return Array.from(new Set(issues))
  }

  function printSchoolProgramme(schoolId: string) {
    const school = programmeBuilder.schools.find((item) => item.id === schoolId)
    if (!school) return
    const groups = builderGroups.filter((entry) => entry.school.id === schoolId).map((entry) => entry.group)
    if (!groups.length) { setProgrammeBuilderMessage('Add at least one group before printing.'); return }

    const escapeHtml = (value: string) => value.replace(/[&<>'"]/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[character] ?? character))
    const activityFill = (code: string) => activities.find((activity) => activity.code.toUpperCase() === code.toUpperCase())?.colour ?? '#f2f4f3'
    const activityTextColour = (code: string) => {
      const hex = activityFill(code).replace('#', '').padEnd(6, '0').slice(0, 6)
      const r = parseInt(hex.slice(0, 2), 16)
      const g = parseInt(hex.slice(2, 4), 16)
      const b = parseInt(hex.slice(4, 6), 16)
      return (r * 299 + g * 587 + b * 114) / 1000 < 145 ? '#ffffff' : '#17211b'
    }
    const printDays = builderDays.filter((dayInfo) => dayInfo.date >= normaliseBuilderDate(school.arrivalDate) && dayInfo.date <= normaliseBuilderDate(school.departureDate))
    const rows = printDays.map((dayInfo) => {
      const activeSessions = BUILDER_SESSIONS.filter((session) => {
        const state = builderSchoolSessionState(school, dayInfo.date, session)
        return state === 'activity' || state === 'arrival'
      })
      return activeSessions.map((session, index) => {
        const state = builderSchoolSessionState(school, dayInfo.date, session)
        const dayCell = index === 0 ? `<th class="day" rowspan="${activeSessions.length}"><span>${escapeHtml(dayInfo.day)}</span></th>` : ''
        if (state === 'arrival') return `<tr>${dayCell}<th class="session">${session}</th><td class="arrival" colspan="${groups.length}">ARRIVAL</td></tr>`
        const groupCells = groups.map((group) => {
          const code = programmeBuilder.assignments[builderAssignmentKey(dayInfo.day, session, group)] ?? ''
          const label = code ? (activities.find((item) => item.code === code)?.name || code) : '—'
          return `<td class="activity" style="background:${activityFill(code)};color:${activityTextColour(code)}"><span>${escapeHtml(label)}</span></td>`
        }).join('')
        return `<tr>${dayCell}<th class="session">${session}</th>${groupCells}</tr>`
      }).join('')
    }).join('<tr class="divider"><td colspan="' + (groups.length + 2) + '"></td></tr>')

    if (!rows.trim()) { setProgrammeBuilderMessage('Build or update this school programme before printing.'); return }
    const logoUrl = new URL(`${import.meta.env.BASE_URL}manor-adventure-logo.png`, window.location.href).toString()
    const dateRange = friendlyProgrammeDateRange(school.arrivalDate, school.departureDate)
    const groupHeaders = groups.map((group) => `<th>G${group}</th>`).join('')
    const popup = window.open('', '_blank', 'width=900,height=1100')
    if (!popup) { setProgrammeBuilderMessage('Allow pop-ups to print the school programme.'); return }
    const documentHtml = `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(school.name || 'School')} programme</title><style>
      @page{size:A4 portrait;margin:9mm}*{box-sizing:border-box;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}html,body{margin:0;padding:0;background:#fff;color:#17211b;font-family:Arial,Helvetica,sans-serif}.sheet{width:100%;min-height:278mm;display:flex;flex-direction:column}.header{display:grid;grid-template-columns:92px 1fr 92px;align-items:center;border-bottom:3px solid #264e3b;padding:0 0 7mm;margin-bottom:5mm}.logo{width:84px;height:58px;object-fit:contain}.heading{text-align:center}.heading h1{font-size:27px;line-height:1;margin:0;color:#244d39;letter-spacing:.3px}.heading h2{font-size:18px;margin:5px 0 0;text-transform:uppercase;letter-spacing:1.1px;color:#111}.heading p{font-size:13px;font-weight:700;margin:5px 0 0;color:#4a5b51}.badge{justify-self:end;width:72px;height:72px;border:2px solid #264e3b;border-radius:50%;display:flex;align-items:center;justify-content:center;text-align:center;font-weight:800;font-size:10px;color:#264e3b;line-height:1.2}table{width:100%;border-collapse:separate;border-spacing:0;table-layout:fixed;font-size:${groups.length > 10 ? '9.5px' : groups.length > 7 ? '10.5px' : '12px'};border:2px solid #24372e;border-radius:5px;overflow:hidden}th,td{border-right:1px solid #334b40;border-bottom:1px solid #334b40;height:${groups.length > 10 ? '27px' : '31px'};padding:3px;text-align:center;font-weight:800;vertical-align:middle}tr:last-child td{border-bottom:0}th:last-child,td:last-child{border-right:0}thead th{background:#244d39;color:#fff;height:28px;font-size:11px;letter-spacing:.4px}.day{width:43px;background:#edf3ef;color:#244d39;padding:0}.day span{writing-mode:vertical-rl;transform:rotate(180deg);display:inline-block;font-size:12px;letter-spacing:.6px}.session{width:39px;background:#f4f6f5;color:#27372f;font-size:11px}.arrival{background:#d95361;color:#fff;font-size:13px;letter-spacing:1.2px}.activity span{display:block;line-height:1.08;overflow-wrap:anywhere}.divider td{height:5px;background:#244d39;padding:0;border-color:#244d39}.departure{background:#244d39;color:#fff;font-size:12px;height:29px;letter-spacing:.2px}.footer{margin-top:auto;padding-top:4mm;text-align:center;border-top:1px solid #aab6af;font-size:9px;color:#56665e}.print-error{margin:30px;text-align:center;font-size:18px}.no-print{position:fixed;top:10px;right:10px;border:0;background:#244d39;color:#fff;padding:10px 14px;border-radius:6px;font-weight:700;cursor:pointer}@media print{.no-print{display:none}.sheet{min-height:auto}}
    </style></head><body><button class="no-print" onclick="window.print()">Print programme</button><main class="sheet"><header class="header"><img class="logo" src="${logoUrl}" alt="Manor Adventure"><div class="heading"><h1>Norfolk Lakes</h1><h2>${escapeHtml(school.name || school.programmeName || 'School Programme')}</h2><p>${escapeHtml(dateRange)}</p></div><div class="badge">SCHOOL<br>PROGRAMME</div></header><table aria-label="School activity programme"><thead><tr><th style="width:43px">DAY</th><th style="width:39px">SES</th>${groupHeaders}</tr></thead><tbody>${rows}<tr><td class="departure" colspan="${groups.length + 2}">Departure after lunch from 1:00–1:30 pm</td></tr></tbody></table><footer class="footer">Manor Adventure · Norfolk Lakes &nbsp;|&nbsp; Activities may change where required by weather or operational conditions.</footer></main><script>
      (function(){let printed=false;function ready(){if(printed)return;printed=true;setTimeout(function(){window.focus();window.print()},350)};const images=Array.from(document.images);if(!images.length){ready();return}let remaining=images.length;images.forEach(function(img){if(img.complete){remaining-=1;if(remaining===0)ready()}else{img.addEventListener('load',function(){remaining-=1;if(remaining===0)ready()},{once:true});img.addEventListener('error',function(){remaining-=1;if(remaining===0)ready()},{once:true})}});setTimeout(ready,2500)})();
    <\/script></body></html>`
    popup.document.open()
    popup.document.write(documentHtml)
    popup.document.close()
    setProgrammeBuilderMessage(`${school.name || 'School'} professional portrait programme opened for printing.`)
  }



  async function downloadSchoolProgrammeExcel(schoolId: string) {
    const school = programmeBuilder.schools.find((item) => item.id === schoolId)
    if (!school) return
    const groups = builderGroups.filter((entry) => entry.school.id === schoolId).map((entry) => entry.group)
    const ExcelJS = (window as Window & { ExcelJS?: any }).ExcelJS
    if (!ExcelJS) { setProgrammeBuilderMessage('The Excel export library did not load. Refresh the page and try again.'); return }
    if (!groups.length) { setProgrammeBuilderMessage('Add at least one group before exporting.'); return }
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('School Programme', { pageSetup: { orientation: 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 1, margins: { left: 0.25, right: 0.25, top: 0.35, bottom: 0.35, header: 0.1, footer: 0.1 } } })
    const lastColumn = groups.length + 2
    sheet.mergeCells(1, 1, 1, lastColumn)
    sheet.getCell(1, 1).value = 'MANOR ADVENTURE · NORFOLK LAKES'
    sheet.getCell(1, 1).font = { bold: true, size: 18, color: { argb: 'FF17324D' } }
    sheet.getCell(1, 1).alignment = { horizontal: 'center', vertical: 'middle' }
    sheet.getRow(1).height = 30
    sheet.mergeCells(2, 1, 2, lastColumn)
    sheet.getCell(2, 1).value = school.name || school.programmeName || 'School Programme'
    sheet.getCell(2, 1).font = { bold: true, size: 15 }
    sheet.getCell(2, 1).alignment = { horizontal: 'center' }
    sheet.mergeCells(3, 1, 3, lastColumn)
    sheet.getCell(3, 1).value = friendlyProgrammeDateRange(school.arrivalDate, school.departureDate)
    sheet.getCell(3, 1).alignment = { horizontal: 'center' }
    const header = sheet.addRow(['DAY', 'SES', ...groups.map((group) => `G${group}`)])
    header.eachCell((cell: any) => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF17324D' } }; cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; cell.alignment = { horizontal: 'center', vertical: 'middle' }; cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } } })
    const printDays = builderDays.filter((dayInfo) => dayInfo.date >= normaliseBuilderDate(school.arrivalDate) && dayInfo.date <= normaliseBuilderDate(school.departureDate))
    for (const dayInfo of printDays) {
      for (const session of BUILDER_SESSIONS) {
        const stateForSchool = builderSchoolSessionState(school, dayInfo.date, session)
        if (stateForSchool !== 'activity' && stateForSchool !== 'arrival') continue
        const values = groups.map((group) => stateForSchool === 'arrival' ? 'ARRIVAL' : (programmeBuilder.assignments[builderAssignmentKey(dayInfo.day, session, group)] ?? ''))
        const row = sheet.addRow([dayInfo.day, session, ...values])
        row.height = 25
        row.eachCell((cell: any, col: number) => {
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
          cell.border = { top: { style: 'thin', color: { argb: 'FF9BA7B1' } }, left: { style: 'thin', color: { argb: 'FF9BA7B1' } }, bottom: { style: 'thin', color: { argb: 'FF9BA7B1' } }, right: { style: 'thin', color: { argb: 'FF9BA7B1' } } }
          if (col > 2) {
            const code = String(cell.value ?? '')
            const colour = code === 'ARRIVAL' ? '#d95361' : (activities.find((activity) => activity.code === code)?.colour ?? '#f3f5f7')
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colour.replace('#', '').toUpperCase()}` } }
            cell.font = { bold: true }
          } else cell.font = { bold: true }
        })
      }
    }
    sheet.columns = [{ width: 14 }, { width: 8 }, ...groups.map(() => ({ width: Math.max(12, Math.min(20, 85 / Math.max(groups.length, 1))) }))]
    sheet.views = [{ state: 'frozen', ySplit: 4, xSplit: 2 }]
    sheet.headerFooter.oddFooter = 'Manor Adventure · Norfolk Lakes · School Programme'
    const buffer = await workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `${(school.name || 'school').replace(/[^a-z0-9_-]+/gi, '-')}-programme.xlsx`; a.click(); URL.revokeObjectURL(url)
    setProgrammeBuilderMessage(`${school.name || 'School'} Excel programme downloaded.`)
  }

  function printCentreProgrammeBuilder() {
    const escapeHtml = (value: string) => value.replace(/[&<>'"]/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[character] ?? character))
    const groups = builderGroups
    if (!groups.length) { setProgrammeBuilderMessage('Add at least one group before exporting.'); return }
    const rows = builderDays.flatMap((dayInfo) => BUILDER_SESSIONS.map((session) => {
      const cells = groups.map(({ group, school }) => {
        const stateForSchool = builderSchoolSessionState(school, dayInfo.date, session)
        const code = programmeBuilder.assignments[builderAssignmentKey(dayInfo.day, session, group)] ?? ''
        const label = stateForSchool === 'arrival' ? 'ARRIVAL' : stateForSchool === 'departed' ? 'DEPARTED' : stateForSchool === 'offsite' ? 'OFF SITE' : code || '—'
        const colour = label === 'ARRIVAL' ? '#d95361' : (activities.find((activity) => activity.code === code)?.colour ?? '#f5f6f7')
        return `<td style="background:${colour}">${escapeHtml(label)}</td>`
      }).join('')
      return `<tr><th>${escapeHtml(dayInfo.day)}</th><th>${session}</th>${cells}</tr>`
    })).join('')
    const headers = groups.map(({ group, school }) => `<th>G${group}<small>${escapeHtml(school.name || 'School')}</small></th>`).join('')
    const popup = window.open('', '_blank', 'width=1400,height=900')
    if (!popup) { setProgrammeBuilderMessage('Allow pop-ups to export the centre programme PDF.'); return }
    popup.document.write(`<!doctype html><html><head><title>Centre Programme</title><style>@page{size:A4 landscape;margin:7mm}*{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}body{font-family:Arial;margin:0;color:#17212b}.head{display:flex;justify-content:space-between;align-items:end;border-bottom:3px solid #2f6f9f;margin-bottom:8px}.head h1{margin:0;font-size:24px;color:#17324d}.head p{margin:2px 0 7px}table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:${groups.length>18?'7px':groups.length>12?'8px':'9px'}th,td{border:1px solid #80909d;padding:4px;text-align:center;font-weight:700;overflow-wrap:anywhere}thead th{background:#17324d;color:white}thead small{display:block;font-size:7px;margin-top:2px}.no-print{position:fixed;right:10px;top:10px;background:#17324d;color:white;border:0;border-radius:6px;padding:9px 12px;font-weight:700}@media print{.no-print{display:none}}</style></head><body><button class="no-print" onclick="window.print()">Save / Print PDF</button><div class="head"><div><h1>Norfolk Lakes Centre Programme</h1><p>${escapeHtml(programmeBuilder.name || 'Programme')}</p></div><strong>${escapeHtml(friendlyProgrammeDateRange(programmeBuilder.startDate, programmeBuilder.endDate))}</strong></div><table><thead><tr><th>DAY</th><th>SES</th>${headers}</tr></thead><tbody>${rows}</tbody></table><script>setTimeout(()=>window.print(),500)<\/script></body></html>`)
    popup.document.close()
    setProgrammeBuilderMessage('Professional centre programme opened for PDF printing.')
  }

  async function downloadCentreProgrammeBuilderExcel() {
    const generated = publishProgrammeDraft(programmeBuilder)
    void downloadPublishedProgrammeExcel(generated)
  }

  function publishProgrammeDraft(draft: ProgrammeBuilderDraft) {
    const days = builderDateRange(draft.startDate, draft.endDate)
    let nextGroup = 1
    const groups = draft.schools.flatMap((school) => Array.from({ length: Math.max(1, school.groups) }, () => ({ school, group: nextGroup++ })))
    const rows: ProgrammeRow[] = []
    for (const dayInfo of days) {
      for (const session of BUILDER_SESSIONS) {
        const arrivalSchools: string[] = []
        const cells = groups.map(({ school, group }) => {
          const state = builderSchoolSessionState(school, dayInfo.date, session)
          if (state === 'arrival' && !arrivalSchools.includes(school.name.trim())) arrivalSchools.push(school.name.trim())
          return {
            group,
            activityCode: state === 'arrival'
              ? school.name.trim()
              : state === 'activity'
                ? (draft.assignments[builderAssignmentKey(dayInfo.day, session, group)] ?? '')
                : '',
          }
        })
        rows.push({
          id: `builder-${dayInfo.day}-${session}`,
          day: dayInfo.day,
          session,
          schoolLabel: arrivalSchools.join(' / '),
          cells,
        })
      }
    }
    const next: ProgrammeImport = {
      title: draft.name.trim() || 'Saved programme', sheetName: 'Programme Builder',
      groupNumbers: groups.map(({ group }) => group), rows, importedAt: new Date().toISOString(),
      sourceFileName: `${draft.name.trim().replace(/[^a-z0-9]+/gi, '-') || 'programme'}-built-in-app.xlsx`,
      startDate: draft.startDate, endDate: draft.endDate,
      schoolDetails: draft.schools.map((school) => ({
        id: school.id,
        schoolName: school.name.trim(),
        programmeName: school.programmeName.trim(),
        purchaseType: school.purchaseType,
        arrivalDate: school.arrivalDate,
        departureDate: school.departureDate,
        notes: school.notes,
        groupNumbers: groups.filter((entry) => entry.school.id === school.id).map((entry) => entry.group),
      })),
    }
    saveProgramme(next, programme ?? undefined)
    ensureWorkingStaffForDays(days.map((entry) => entry.day))
    setProgrammeBuilderMessage('Programme loaded into Programme and Daily Staffing.')
    setImportMessage(`Loaded ${next.title} from the saved programme library.`)
    setPage('programme')
    return next
  }

  function publishProgrammeBuilder() {
    const issues = builderValidation()
    if (issues.length) { setProgrammeBuilderMessage(`Cannot publish: ${issues[0]}`); return }
    publishProgrammeDraft(programmeBuilder)
  }

  function saveProgramme(next: ProgrammeImport, previous?: ProgrammeImport) {
    const nextHistory = previous ? [previous, ...history].slice(0, 12) : history
    if (previous) archiveSnapshot(previous, true)
    const migrated = migrateAssignments(previous ?? null, next, assignments)
    setAssignments(migrated)
    localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(migrated))
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


  async function importAppProgramme(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    try {
      setImportMessage('Opening app programme…')
      const data = await file.arrayBuffer()
      const workbook = XLSX.read(data, { type: 'array' })
      const stateSheet = workbook.Sheets['_ACM_DATA']
      if (!stateSheet) throw new Error('This is not an Adventure Centre Manager app programme. Use Upload Excel programme for ordinary spreadsheets.')
      const raw = String(stateSheet.A1?.v ?? '')
      const payload = JSON.parse(raw) as { programme?: ProgrammeImport; draft?: ProgrammeBuilderDraft }
      if (!payload.programme?.rows?.length) throw new Error('The app programme data is incomplete.')
      const restored = { ...payload.programme, importedAt: new Date().toISOString(), sourceFileName: file.name }
      saveProgramme(restored, programme ?? undefined)
      if (payload.draft) {
        const normalised: ProgrammeBuilderDraft = { ...payload.draft, manualLocks: payload.draft.manualLocks ?? {}, schools: (payload.draft.schools ?? []).map((school) => ({ ...school, programmeName: school.programmeName ?? school.name ?? '', purchaseType: school.purchaseType ?? payload.draft!.purchaseType ?? 'normal', arrivalDate: school.arrivalDate ?? payload.draft!.startDate ?? '', departureDate: school.departureDate ?? payload.draft!.endDate ?? '', notes: school.notes ?? payload.draft!.notes ?? '', requestedActivities: school.requestedActivities ?? [], backupOption1: school.backupOption1 ?? '', backupOption2: school.backupOption2 ?? '', locked: school.locked ?? false })) }
        setProgrammeBuilder(normalised)
        localStorage.setItem(PROGRAMME_BUILDER_KEY, JSON.stringify(normalised))
      }
      ensureWorkingStaffForDays(Array.from(new Set(restored.rows.map((row) => row.day))))
      setImportMessage(`Opened ${restored.title || file.name} with its editable app data.`)
      setPage('programme')
    } catch (error) {
      setImportMessage(error instanceof Error ? error.message : 'The app programme could not be opened.')
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

  function programmeGroupSchool(source: ProgrammeImport, group: number) {
    return source.schoolDetails?.find((school) => school.groupNumbers?.includes(group))?.schoolName ?? ''
  }

  function programmeCellDisplay(row: ProgrammeRow, group: number) {
    const value = row.cells.find((cell) => cell.group === group)?.activityCode ?? ''
    if (!value) return '—'
    const schoolName = programmeGroupSchool(programme!, group)
    if (row.session === '3' && schoolName && value.toLowerCase() === schoolName.toLowerCase()) return schoolName
    return value
  }

  async function downloadPublishedProgrammeExcel(sourceProgramme?: ProgrammeImport) {
    const exportProgramme: ProgrammeImport | null = sourceProgramme ?? programme
    if (!exportProgramme) return
    const ExcelJS = (window as Window & { ExcelJS?: any }).ExcelJS
    if (!ExcelJS) {
      setImportMessage('The Excel export library did not load. Refresh the page and try again.')
      return
    }

    const title = exportProgramme.title || exportProgramme.sourceFileName.replace(/\.xlsx$/i, '') || 'Centre Programme'
    const mergedRows = Array.from(exportProgramme.rows.reduce((map, row) => {
      const key = `${row.day}|${row.session}`
      const current = map.get(key)
      if (!current) map.set(key, { ...row, cells: row.cells.map((cell) => ({ ...cell })) })
      else {
        const cells = new Map(current.cells.map((cell) => [cell.group, cell]))
        row.cells.forEach((cell) => {
          const existing = cells.get(cell.group)
          if (!existing?.activityCode || cell.activityCode) cells.set(cell.group, { ...cell })
        })
        map.set(key, { ...current, cells: Array.from(cells.values()) })
      }
      return map
    }, new Map<string, ProgrammeRow>()).values()).sort((a, b) => weekdayRank(a.day) - weekdayRank(b.day) || Number(a.session) - Number(b.session))

    const workbook = new ExcelJS.Workbook()
    workbook.creator = 'Adventure Centre Manager'
    workbook.company = 'Manor Adventure – Norfolk Lakes'
    workbook.title = title
    workbook.subject = 'Centre Programme'
    workbook.created = new Date()

    const sheet = workbook.addWorksheet('Centre Programme', {
      views: [{ state: 'frozen', xSplit: 2, ySplit: 6, topLeftCell: 'C7', activeCell: 'C7' }],
      properties: { defaultRowHeight: 20, showGridLines: false },
      pageSetup: {
        paperSize: 9,
        orientation: 'landscape',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        horizontalCentered: true,
        verticalCentered: false,
        margins: { left: 0.25, right: 0.25, top: 0.35, bottom: 0.35, header: 0.15, footer: 0.15 },
      },
      headerFooter: {
        oddFooter: '&LManor Adventure · Norfolk Lakes&CPage &P of &N&RGenerated &D',
      },
    })

    const totalColumns = Math.max(3, exportProgramme.groupNumbers.length + 2)
    const lastColumn = sheet.getColumn(totalColumns).letter
    const darkGreen = '173F37'
    const midGreen = '2D6657'
    const paleGreen = 'E7F0EC'
    const borderDark = '334B43'
    const schoolPalette = ['DCEFE8', 'E7E1F3', 'F8E2D8', 'DCE8F7', 'F5E8BA', 'E1ECD2']

    sheet.mergeCells(`A1:${lastColumn}1`)
    const titleCell = sheet.getCell('A1')
    titleCell.value = 'NORFOLK LAKES — CENTRE PROGRAMME'
    titleCell.font = { name: 'Aptos Display', size: 20, bold: true, color: { argb: 'FFFFFFFF' } }
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${darkGreen}` } }
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
    sheet.getRow(1).height = 32

    sheet.mergeCells(`A2:${lastColumn}2`)
    const programmeTitleCell = sheet.getCell('A2')
    programmeTitleCell.value = title.toUpperCase()
    programmeTitleCell.font = { name: 'Aptos', size: 14, bold: true, color: { argb: `FF${darkGreen}` } }
    programmeTitleCell.alignment = { horizontal: 'center', vertical: 'middle' }
    sheet.getRow(2).height = 23

    sheet.mergeCells(`A3:${lastColumn}3`)
    const dateCell = sheet.getCell('A3')
    dateCell.value = friendlyProgrammeDateRange(exportProgramme.startDate ?? '', exportProgramme.endDate ?? '')
    dateCell.font = { name: 'Aptos', size: 10, bold: true, color: { argb: 'FF53645D' } }
    dateCell.alignment = { horizontal: 'center', vertical: 'middle' }
    sheet.getRow(3).height = 18
    sheet.getRow(4).height = 7

    sheet.mergeCells('A5:A6')
    sheet.mergeCells('B5:B6')
    sheet.getCell('A5').value = 'DAY'
    sheet.getCell('B5').value = 'SES'

    const schoolBlocks: Array<{ name: string; start: number; end: number; colour: string }> = []
    let blockStart = 3
    let previousSchool = ''
    exportProgramme.groupNumbers.forEach((group, index) => {
      const school = programmeGroupSchool(exportProgramme, group) || 'Unassigned'
      const column = index + 3
      if (school !== previousSchool) {
        if (schoolBlocks.length) schoolBlocks[schoolBlocks.length - 1].end = column - 1
        schoolBlocks.push({ name: school, start: column, end: column, colour: schoolPalette[schoolBlocks.length % schoolPalette.length] })
        blockStart = column
        previousSchool = school
      } else {
        schoolBlocks[schoolBlocks.length - 1].end = column
      }
      sheet.getCell(6, column).value = `G${group}`
    })
    void blockStart

    schoolBlocks.forEach((block) => {
      if (block.end > block.start) sheet.mergeCells(5, block.start, 5, block.end)
      const cell = sheet.getCell(5, block.start)
      cell.value = block.name.toUpperCase()
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${block.colour}` } }
      cell.font = { name: 'Aptos', size: 10, bold: true, color: { argb: `FF${darkGreen}` } }
      cell.alignment = { horizontal: 'center', vertical: 'middle' }
    })

    for (let column = 1; column <= totalColumns; column += 1) {
      const headerCell = sheet.getCell(6, column)
      if (column <= 2) {
        const mergedHeader = sheet.getCell(5, column)
        mergedHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${darkGreen}` } }
        mergedHeader.font = { name: 'Aptos', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }
        mergedHeader.alignment = { horizontal: 'center', vertical: 'middle' }
      } else {
        headerCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${midGreen}` } }
        headerCell.font = { name: 'Aptos', size: 9, bold: true, color: { argb: 'FFFFFFFF' } }
        headerCell.alignment = { horizontal: 'center', vertical: 'middle' }
      }
    }
    sheet.getRow(5).height = 23
    sheet.getRow(6).height = 21

    const colourForCode = (code: string) => {
      const saved = activities.find((activity) => activity.code.toUpperCase() === code.toUpperCase())?.colour
      return (saved || '#F2F4F3').replace('#', '').toUpperCase().padEnd(6, '0').slice(0, 6)
    }
    const readableText = (hex: string) => {
      const r = parseInt(hex.slice(0, 2), 16)
      const g = parseInt(hex.slice(2, 4), 16)
      const b = parseInt(hex.slice(4, 6), 16)
      return (r * 299 + g * 587 + b * 114) / 1000 < 145 ? 'FFFFFF' : '16211C'
    }
    const displayCode = (value: string) => {
      const schoolName = (exportProgramme.schoolDetails ?? []).find((school) => school.schoolName.toLowerCase() === value.toLowerCase())?.schoolName
      if (schoolName) return 'ARRIVAL'
      return value === '—' ? '' : value
    }

    let outputRow = 7
    let currentDay = ''
    let dayStartRow = outputRow
    const dayRanges: Array<{ start: number; end: number; day: string }> = []

    mergedRows.forEach((row, index) => {
      if (currentDay && row.day !== currentDay) {
        dayRanges.push({ start: dayStartRow, end: outputRow - 1, day: currentDay })
        dayStartRow = outputRow
      }
      currentDay = row.day
      const excelRow = sheet.getRow(outputRow)
      excelRow.height = 22
      excelRow.getCell(1).value = row.day
      excelRow.getCell(2).value = row.session

      exportProgramme.groupNumbers.forEach((group, groupIndex) => {
        const rawValue = programmeCellDisplay(row, group)
        const value = displayCode(rawValue)
        const cell = excelRow.getCell(groupIndex + 3)
        cell.value = value
        const isArrival = value === 'ARRIVAL'
        const fill = isArrival ? 'D85261' : value ? colourForCode(value) : 'F4F6F5'
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${fill}` } }
        cell.font = { name: 'Aptos Narrow', size: 9, bold: true, color: { argb: `FF${isArrival ? 'FFFFFF' : readableText(fill)}` } }
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true, shrinkToFit: true }
      })
      outputRow += 1
      if (index === mergedRows.length - 1 && currentDay) dayRanges.push({ start: dayStartRow, end: outputRow - 1, day: currentDay })
    })

    dayRanges.forEach(({ start, end, day }) => {
      if (end > start) sheet.mergeCells(start, 1, end, 1)
      const dayCell = sheet.getCell(start, 1)
      dayCell.value = day.toUpperCase()
      dayCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${paleGreen}` } }
      dayCell.font = { name: 'Aptos', size: 9, bold: true, color: { argb: `FF${darkGreen}` } }
      dayCell.alignment = { horizontal: 'center', vertical: 'middle', textRotation: 90 }
      for (let column = 1; column <= totalColumns; column += 1) {
        sheet.getCell(end, column).border = {
          ...(sheet.getCell(end, column).border || {}),
          bottom: { style: 'medium', color: { argb: `FF${darkGreen}` } },
        }
      }
    })

    for (let row = 5; row < outputRow; row += 1) {
      for (let column = 1; column <= totalColumns; column += 1) {
        const cell = sheet.getCell(row, column)
        const schoolBoundary = schoolBlocks.some((block) => block.end === column)
        cell.border = {
          ...(cell.border || {}),
          top: cell.border?.top || { style: 'thin', color: { argb: `FF${borderDark}` } },
          bottom: cell.border?.bottom || { style: 'thin', color: { argb: `FF${borderDark}` } },
          left: cell.border?.left || { style: 'thin', color: { argb: `FF${borderDark}` } },
          right: { style: schoolBoundary ? 'medium' : 'thin', color: { argb: `FF${borderDark}` } },
        }
      }
    }

    for (let row = 7; row < outputRow; row += 1) {
      const sessionCell = sheet.getCell(row, 2)
      sessionCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F8F7' } }
      sessionCell.font = { name: 'Aptos', size: 9, bold: true, color: { argb: `FF${darkGreen}` } }
      sessionCell.alignment = { horizontal: 'center', vertical: 'middle' }
    }

    sheet.getColumn(1).width = 6
    sheet.getColumn(2).width = 5
    exportProgramme.groupNumbers.forEach((_, index) => { sheet.getColumn(index + 3).width = 10.5 })
    sheet.autoFilter = { from: { row: 6, column: 1 }, to: { row: outputRow - 1, column: totalColumns } }
    sheet.pageSetup.printArea = `A1:${lastColumn}${outputRow - 1}`
    sheet.pageSetup.printTitlesRow = '1:6'

    const legend = workbook.addWorksheet('Activity Legend', { properties: { showGridLines: false } })
    legend.columns = [{ width: 15 }, { width: 32 }, { width: 18 }, { width: 20 }]
    legend.mergeCells('A1:D1')
    legend.getCell('A1').value = 'ACTIVITY COLOUR LEGEND'
    legend.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${darkGreen}` } }
    legend.getCell('A1').font = { name: 'Aptos Display', size: 16, bold: true, color: { argb: 'FFFFFFFF' } }
    legend.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }
    legend.getRow(1).height = 28
    ;['Code', 'Activity', 'Maximum groups at once', 'Equipment quantity'].forEach((label, index) => {
      const cell = legend.getCell(3, index + 1)
      cell.value = label
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${midGreen}` } }
      cell.font = { name: 'Aptos', bold: true, color: { argb: 'FFFFFFFF' } }
      cell.alignment = { horizontal: 'center', vertical: 'middle' }
    })
    enabledActivities.forEach((activity, index) => {
      const row = legend.getRow(index + 4)
      const fill = (activity.colour || '#F2F4F3').replace('#', '').toUpperCase()
      row.values = [activity.code, activity.name, activity.capacity ?? 1, activity.equipmentQuantity ?? 0]
      row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${fill}` } }
      row.getCell(1).font = { name: 'Aptos', bold: true, color: { argb: `FF${readableText(fill)}` } }
      row.eachCell((cell: any) => {
        cell.alignment = { horizontal: 'center', vertical: 'middle' }
        cell.border = { top: { style: 'thin', color: { argb: 'FFB7C2BD' } }, bottom: { style: 'thin', color: { argb: 'FFB7C2BD' } }, left: { style: 'thin', color: { argb: 'FFB7C2BD' } }, right: { style: 'thin', color: { argb: 'FFB7C2BD' } } }
      })
    })

    const appState = workbook.addWorksheet('_ACM_DATA', { state: 'veryHidden' })
    appState.getCell('A1').value = JSON.stringify({ exportProgramme, draft: programmeBuilder, exportedAt: new Date().toISOString(), format: 'ACM_APP_PROGRAMME_V1' })

    try {
      const logoResponse = await fetch(`${import.meta.env.BASE_URL}manor-adventure-logo.png`)
      if (logoResponse.ok) {
        const logoBuffer = await logoResponse.arrayBuffer()
        const logoId = workbook.addImage({ buffer: logoBuffer, extension: 'png' })
        sheet.addImage(logoId, { tl: { col: 0.1, row: 0.1 }, ext: { width: 90, height: 42 } })
      }
    } catch {
      // Branding text remains visible if the logo cannot be loaded.
    }

    const bytes = await workbook.xlsx.writeBuffer()
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const link = document.createElement('a')
    link.href = URL.createObjectURL(blob)
    const safeName = title.replace(/[^a-z0-9 _-]+/gi, '').trim() || 'Centre Programme'
    link.download = `${safeName}-professional.xlsx`
    document.body.appendChild(link)
    link.click()
    link.remove()
    URL.revokeObjectURL(link.href)
    setImportMessage('Downloaded the professional centre exportProgramme workbook using your saved activity colours.')
  }

  function schoolProgrammeFromArchive(source: ProgrammeImport, schoolName: string) {
    const school = (source.schoolDetails ?? []).find((item) => item.schoolName === schoolName)
    if (!school) return source
    const groups = new Set(school.groupNumbers ?? [])
    return {
      ...source,
      title: school.programmeName || `${school.schoolName} Programme`,
      sourceFileName: `${school.schoolName}-programme.xlsx`,
      groupNumbers: source.groupNumbers.filter((group) => groups.has(group)),
      rows: source.rows.map((row) => ({ ...row, cells: row.cells.filter((cell) => groups.has(cell.group)) })),
      schoolDetails: [school],
    }
  }

  function printArchivedProgramme(source: ProgrammeImport, schoolName?: string) {
    const exportProgramme = schoolName ? schoolProgrammeFromArchive(source, schoolName) : source
    const groups = exportProgramme.groupNumbers
    const headers = groups.map((group) => `<th>G${group}<small>${escapeHtml(exportProgramme.schoolDetails?.find((school) => school.groupNumbers?.includes(group))?.schoolName ?? '')}</small></th>`).join('')
    const rows = exportProgramme.rows.map((row) => `<tr><th>${escapeHtml(row.day)}</th><th>${escapeHtml(row.session)}</th>${groups.map((group) => `<td>${escapeHtml(row.cells.find((cell) => cell.group === group)?.activityCode ?? '')}</td>`).join('')}</tr>`).join('')
    const popup = window.open('', '_blank')
    if (!popup) { setImportMessage('Allow pop-ups to download the programme PDF.'); return }
    popup.document.write(`<!doctype html><html><head><title>${escapeHtml(exportProgramme.title)}</title><style>@page{size:A4 landscape;margin:8mm}*{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}body{font-family:Arial;margin:0;color:#17212b}.head{display:flex;justify-content:space-between;align-items:end;border-bottom:3px solid #2f6f9f;margin-bottom:10px}.head h1{margin:0;font-size:24px;color:#17324d}.head p{margin:4px 0 8px}table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:${groups.length > 18 ? '7px' : groups.length > 12 ? '8px' : '9px'}th,td{border:1px solid #80909d;padding:4px;text-align:center;font-weight:700;overflow-wrap:anywhere}thead th{background:#17324d;color:white}thead small{display:block;font-size:7px;margin-top:2px}.no-print{position:fixed;right:10px;top:10px;background:#17324d;color:white;border:0;border-radius:6px;padding:9px 12px;font-weight:700}@media print{.no-print{display:none}}</style></head><body><button class="no-print" onclick="window.print()">Save / Print PDF</button><div class="head"><div><h1>${escapeHtml(exportProgramme.title)}</h1><p>${schoolName ? escapeHtml(schoolName) : 'Complete Centre Programme'}</p></div><strong>${escapeHtml(friendlyProgrammeDateRange(exportProgramme.startDate ?? '', exportProgramme.endDate ?? ''))}</strong></div><table><thead><tr><th>DAY</th><th>SES</th>${headers}</tr></thead><tbody>${rows}</tbody></table><script>setTimeout(()=>window.print(),500)<\/script></body></html>`)
    popup.document.close()
  }

  function deleteArchivedProgramme(archiveId: string) {
    if (!window.confirm('Delete this programme from the archive?')) return
    const next = programmeArchive.filter((item) => item.archiveId !== archiveId)
    setProgrammeArchive(next)
    localStorage.setItem(PROGRAMME_ARCHIVE_KEY, JSON.stringify(next))
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

  function nextStaffCode() {
    const all = [...staff, ...formerStaff.map((x) => x.member), ...loanHistory.map((x) => x.member)]
    const max = all.reduce((value, member) => Math.max(value, Number(member.staffCode?.replace(/\D/g, '') || 0)), 0)
    return `NL${String(max + 1).padStart(4, '0')}`
  }

  function addTimeline(memberId: string, event: string, date = dateKey(new Date())) {
    setStaffTimeline((current) => {
      const next = { ...current, [memberId]: [...(current[memberId] ?? []), { date, event }] }
      localStorage.setItem(STAFF_TIMELINE_KEY, JSON.stringify(next))
      return next
    })
  }

  function saveFormer(next: ArchivedStaff[]) { setFormerStaff(next); localStorage.setItem(FORMER_STAFF_KEY, JSON.stringify(next)) }
  function saveLoanHistory(next: ArchivedStaff[]) { setLoanHistory(next); localStorage.setItem(LOAN_HISTORY_KEY, JSON.stringify(next)) }

  function removeFromActiveStaff(staffId: string) {
    const nextStaff = staff.filter((item) => item.id !== staffId)
    const nextAssignments = Object.fromEntries(Object.entries(assignments).filter(([, assignedId]) => assignedId !== staffId))
    const nextSickness = Object.fromEntries(Object.entries(sicknessByDay).map(([day, ids]) => [day, ids.filter((id) => id !== staffId)]))
    setStaff(nextStaff); setAssignments(nextAssignments); setSicknessByDay(nextSickness)
    localStorage.setItem(STAFF_KEY, JSON.stringify(nextStaff)); localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(nextAssignments)); localStorage.setItem(SICKNESS_KEY, JSON.stringify(nextSickness))
  }

  function markLeftCompany(staffId: string) {
    const member = staff.find((item) => item.id === staffId); if (!member) return
    const leavingDate = window.prompt('Leaving date (YYYY-MM-DD)', dateKey(new Date())); if (!leavingDate) return
    const notes = window.prompt('Optional leaving note', '') ?? ''
    if (!window.confirm(`Move ${member.name} to Former Staff?`)) return
    saveFormer([...formerStaff, { member: { ...member, employmentType: 'permanent' }, archivedAt: new Date().toISOString(), endDate: leavingDate, notes, archiveType: 'former' }])
    removeFromActiveStaff(staffId); addTimeline(member.id, `Left company${notes ? ` — ${notes}` : ''}`, leavingDate)
    setImportMessage(`${member.name} was moved to Former Staff.`)
  }

  function endLoan(staffId: string) {
    const member = staff.find((item) => item.id === staffId); if (!member) return
    const endDate = window.prompt('Loan end date (YYYY-MM-DD)', dateKey(new Date())); if (!endDate) return
    const notes = window.prompt('Optional loan note', '') ?? ''
    const prior = loanHistory.find((x) => x.member.id === member.id)
    const period = { startDate: member.startDate ?? endDate, endDate, notes }
    const next = prior ? loanHistory.map((x) => x.member.id === member.id ? { ...x, member, endDate, archivedAt: new Date().toISOString(), loanPeriods: [...(x.loanPeriods ?? []), period] } : x) : [...loanHistory, { member, archivedAt: new Date().toISOString(), endDate, notes, archiveType: 'loan' as const, loanPeriods: [period] }]
    saveLoanHistory(next); removeFromActiveStaff(staffId); addTimeline(member.id, `Loan ended${notes ? ` — ${notes}` : ''}`, endDate)
    setImportMessage(`${member.name}'s loan was ended and saved in Loan Staff History.`)
  }

  function reinstateFormer(record: ArchivedStaff) {
    const startDate = window.prompt('Reinstatement date (YYYY-MM-DD)', dateKey(new Date())); if (!startDate) return
    const member = { ...record.member, startDate, employmentType: 'permanent' as const }
    const next = [...staff, member]; setStaff(next); localStorage.setItem(STAFF_KEY, JSON.stringify(next))
    saveFormer(formerStaff.filter((x) => x.member.id !== record.member.id)); addTimeline(member.id, 'Reinstated at centre', startDate)
  }

  function reactivateLoan(record: ArchivedStaff) {
    const startDate = window.prompt('New loan start date (YYYY-MM-DD)', dateKey(new Date())); if (!startDate) return
    const endDate = window.prompt('Expected loan end date (YYYY-MM-DD)', '') ?? ''
    const member = { ...record.member, startDate, loanEndDate: endDate, employmentType: 'loan' as const }
    const next = [...staff, member]; setStaff(next); localStorage.setItem(STAFF_KEY, JSON.stringify(next))
    saveLoanHistory(loanHistory.filter((x) => x.member.id !== record.member.id)); addTimeline(member.id, 'Reactivated as loan staff', startDate)
  }

  function convertLoanToPermanent(record: ArchivedStaff) {
    const startDate = window.prompt('Permanent employment start date (YYYY-MM-DD)', dateKey(new Date())); if (!startDate) return
    const role = (window.prompt('Role: staff, teamLeader, activityManager or centreManager', resolvedRole(record.member)) || resolvedRole(record.member)) as StaffRole
    const member = { ...record.member, startDate, role, teamLeader: role === 'teamLeader', employmentType: 'permanent' as const, loanEndDate: undefined }
    const next = [...staff, member]; setStaff(next); localStorage.setItem(STAFF_KEY, JSON.stringify(next))
    saveLoanHistory(loanHistory.filter((x) => x.member.id !== record.member.id)); addTimeline(member.id, `Added permanently as ${roleLabel(role)}`, startDate)
  }

  function payrollSync() {
    const now = new Date().toISOString(); localStorage.setItem(PAYROLL_SYNC_KEY, now); setPayrollSyncAt(now)
    const permanentCount = staff.filter((member) => member.employmentType !== 'loan').length
    setImportMessage(`Payroll synced with ${permanentCount} permanent staff. Existing template order will be retained and new staff will be added to the bottom of their role group.`)
  }

  async function makeAdminAccount() {
    const email = adminAccountEmail.trim().toLowerCase()
    if (!email) return
    setAdminAccountBusy(true)
    const { data, error } = await supabase
      .from('profiles')
      .update({ role: 'admin' })
      .eq('email', email)
      .select('id,email')

    setAdminAccountBusy(false)
    if (error) {
      setImportMessage(`Could not give admin access: ${error.message}`)
      return
    }
    if (!data?.length) {
      setImportMessage('No account was found with that email. Ask the admin person to create an account on the login screen first, then try again.')
      return
    }
    setAdminAccountEmail('')
    setImportMessage(`${email} now has access to Admin → Programme Files only.`)
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
      staffCode: nextStaffCode(),
      employmentType: addingLoanStaff ? 'loan' : 'permanent',
      startDate: newStaffStartDate,
      loanEndDate: addingLoanStaff ? newLoanEndDate : undefined,
    }

    const next = [...staff, nextMember]
    setStaff(next)
    localStorage.setItem(STAFF_KEY, JSON.stringify(next))
    setNewStaffName('')
    setNewStaffRole('staff')
    setNewStaffQualifications([])
    setShowAddStaff(false)
    addTimeline(nextMember.id, addingLoanStaff ? 'Joined as loan staff' : 'Joined centre', newStaffStartDate)
    setAddingLoanStaff(false)
    setNewLoanEndDate('')
    setImportMessage(`${trimmedName} was added to the staff platform${nextMember.employmentType === 'loan' ? ' as loan staff' : ''}.`)
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

    const nextWaterSupport = Object.fromEntries(
      Object.entries(waterSupportAssignments).filter(([key]) => !key.startsWith(`${day}::`)),
    )
    setAssignments(nextAssignments)
    setArrivalAssignments(nextArrivalAssignments)
    setWaterSupportAssignments(nextWaterSupport)
    localStorage.setItem(WATER_SUPPORT_KEY, JSON.stringify(nextWaterSupport))
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


  type QualificationShortage = {
    session: string
    activityCode: string
    activityName: string
    required: number
    covered: number
    shortfall: number
  }

  function qualificationShortagesForDay(day: string): QualificationShortage[] {
    if (!programme || !day) return []
    const workingIds = new Set(workingByDay[day] ?? staff.map((member) => member.id))
    const unavailable = new Set<string>()
    const available = staff.filter((member) => workingIds.has(member.id) && !unavailable.has(member.id))
    const shortages: QualificationShortage[] = []

    const sessions = Array.from(new Set(programme.rows.filter((row) => row.day === day).map((row) => row.session)))
    sessions.forEach((session) => {
      const blockedByArrivals = arrivalStaffForDaySession(day, session)
      const sessionUnavailable = unavailableStaffIdsForSession(day, session)
      const sessionStaff = available.filter((member) => !sessionUnavailable.has(member.id) && !blockedByArrivals.has(member.id))
      const slots = staffingClustersForRows(programme.rows.filter((row) => row.day === day && row.session === session))
        .flatMap((cluster) => Array.from({ length: cluster.required }, (_, index) => ({
          id: `${cluster.key}::slot-${index}`,
          activityCode: cluster.activityCode,
        })))
        .sort((a, b) => {
          const aCount = sessionStaff.filter((member) => qualificationIsValid(member, a.activityCode)).length
          const bCount = sessionStaff.filter((member) => qualificationIsValid(member, b.activityCode)).length
          return aCount - bCount || a.activityCode.localeCompare(b.activityCode)
        })

      const matchedSlotByStaff = new Map<string, string>()
      const matchedStaffBySlot = new Map<string, string>()
      const slotById = new Map(slots.map((slot) => [slot.id, slot]))

      const tryMatch = (slotId: string, visited: Set<string>): boolean => {
        const slot = slotById.get(slotId)
        if (!slot) return false
        const candidates = sessionStaff
          .filter((member) => qualificationIsValid(member, slot.activityCode))
          .sort((a, b) => rolePriority(resolvedRole(a)) - rolePriority(resolvedRole(b)) || a.name.localeCompare(b.name))
        for (const member of candidates) {
          if (visited.has(member.id)) continue
          visited.add(member.id)
          const existingSlot = matchedSlotByStaff.get(member.id)
          if (!existingSlot || tryMatch(existingSlot, visited)) {
            matchedSlotByStaff.set(member.id, slotId)
            matchedStaffBySlot.set(slotId, member.id)
            return true
          }
        }
        return false
      }

      slots.forEach((slot) => tryMatch(slot.id, new Set()))
      const requiredByCode = new Map<string, number>()
      const coveredByCode = new Map<string, number>()
      slots.forEach((slot) => requiredByCode.set(slot.activityCode, (requiredByCode.get(slot.activityCode) ?? 0) + 1))
      matchedStaffBySlot.forEach((_, slotId) => {
        const code = slotById.get(slotId)?.activityCode
        if (code) coveredByCode.set(code, (coveredByCode.get(code) ?? 0) + 1)
      })
      requiredByCode.forEach((required, code) => {
        const covered = coveredByCode.get(code) ?? 0
        if (covered < required) shortages.push({
          session,
          activityCode: code,
          activityName: activityName(code),
          required,
          covered,
          shortfall: required - covered,
        })
      })
    })

    return shortages.sort((a, b) => Number(a.session) - Number(b.session) || a.activityName.localeCompare(b.activityName))
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
    if (!programme) {
      window.alert('No programme is loaded. Load or build a programme before building the rota.')
      return
    }

    let next: StaffingAssignment = { ...assignments }
    const workload = new Map<string, number>()
    const unresolved: string[] = []
    const sortedRows = [...programme.rows].sort((a,b) => a.day.localeCompare(b.day) || Number(a.session)-Number(b.session))

    for (const row of sortedRows) {
      const workingIds = new Set(workingByDay[row.day] ?? staff.map((m) => m.id))
      const unavailableIds = unavailableStaffIdsForSession(row.day, row.session)
      const arrivalIds = arrivalStaffForDaySession(row.day, row.session)

      for (const cell of activityCellsForRow(row)) {
        const key = cellKey(row.id, cell.group)
        if (next[key]) {
          workload.set(next[key], (workload.get(next[key]) ?? 0) + 1)
          continue
        }

        const candidates = staff
          .filter((member) => workingIds.has(member.id))
          .filter((member) => !unavailableIds.has(member.id))
          .filter((member) => qualificationIsValid(member, cell.activityCode))
          .filter((member) => !arrivalIds.has(member.id))
          .filter((member) => !sortedRows.some((other) =>
            other.day === row.day &&
            other.session === row.session &&
            other.cells.some((otherCell) => next[cellKey(other.id, otherCell.group)] === member.id),
          ))
          .sort((a,b) =>
            (rolePriority(resolvedRole(a)) - rolePriority(resolvedRole(b))) ||
            ((workload.get(a.id) ?? 0) - (workload.get(b.id) ?? 0)) ||
            a.name.localeCompare(b.name),
          )

        if (candidates[0]) {
          next[key] = candidates[0].id
          workload.set(candidates[0].id, (workload.get(candidates[0].id) ?? 0) + 1)
        } else {
          unresolved.push(`${row.day}, Session ${row.session}, Group ${cell.group}: ${activityName(cell.activityCode)}`)
        }
      }
    }

    const faultLines: string[] = []
    if (unresolved.length) {
      faultLines.push(`${unresolved.length} activity assignment${unresolved.length === 1 ? '' : 's'} could not be filled.`)
      unresolved.slice(0, 8).forEach((line) => faultLines.push(`• ${line}`))
      if (unresolved.length > 8) faultLines.push(`• Plus ${unresolved.length - 8} more.`)
    }

    const qualificationDays = programmeDays.filter((day) => qualificationShortagesForDay(day).length > 0)
    if (qualificationDays.length) faultLines.push(`Qualification/sign-off shortages detected on: ${qualificationDays.join(', ')}.`)

    const message = faultLines.length
      ? `Faults were found before building the rota:

${faultLines.join('\n')}

Build the rota anyway?`
      : 'No faults were found. Build the entire rota now?'

    if (!window.confirm(message)) return

    setAssignments(next)
    localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(next))
    setImportMessage(faultLines.length
      ? `Rota built with ${unresolved.length} unfilled activity assignment${unresolved.length === 1 ? '' : 's'}. Review the warnings in Daily Staffing.`
      : 'Entire rota built successfully using availability, qualifications, workload balancing and conflict prevention.')
  }

  function isWaterActivity(code: string) {
    const activity = activities.find((item) => item.code === code)
    const text = normaliseActivityText(`${code} ${activity?.name ?? ''}`)
    return ['water', 'sailing', 'kayak', 'canoe', 'paddle', 'raft', 'sup', 'windsurf'].some((term) => text.includes(term))
  }

  function waterDiscipline(code: string): 'canoe' | 'kayak' | null {
    const activity = activities.find((item) => item.code === code)
    const text = normaliseActivityText(`${code} ${activity?.name ?? ''}`)
    if (text.includes('canoe')) return 'canoe'
    if (text.includes('kayak')) return 'kayak'
    return null
  }

  function waterSupportKey(day: string, session: string, discipline: 'canoe' | 'kayak') {
    return `${day}::${session}::${discipline}`
  }

  function leadQualificationCode(discipline: 'canoe' | 'kayak') {
    return discipline === 'canoe' ? 'CANOE LEAD' : 'KAYAK LEAD'
  }

  function qualifiedWaterLeads(discipline: 'canoe' | 'kayak') {
    const code = leadQualificationCode(discipline)
    return staff.filter((member) => member.qualifications.includes(code))
  }

  function disciplineAssignments(day: string, session: string, discipline: 'canoe' | 'kayak', currentAssignments = assignments) {
    if (!programme) return [] as {staffId:string;group:number}[]
    const result: {staffId:string;group:number}[] = []
    programme.rows.filter((row) => row.day === day && row.session === session).forEach((row) => {
      activityCellsForRow(row).forEach((cell) => {
        if (waterDiscipline(cell.activityCode) !== discipline) return
        const staffId = currentAssignments[cellKey(row.id, cell.group)]
        if (staffId) result.push({ staffId, group: cell.group })
      })
    })
    return result
  }

  function waterSupportNeedsForDay(day: string) {
    if (!programme) return []
    const counts = new Map<string, { session: string; discipline: 'canoe' | 'kayak'; groups: number }>()
    programme.rows.filter((row) => row.day === day).forEach((row) => {
      activityCellsForRow(row).forEach((cell) => {
        const discipline = waterDiscipline(cell.activityCode)
        if (!discipline) return
        const key = `${row.session}::${discipline}`
        const current = counts.get(key) ?? { session: row.session, discipline, groups: 0 }
        current.groups += 1
        counts.set(key, current)
      })
    })
    return Array.from(counts.values()).filter((item) => item.groups >= 2)
      .sort((a,b) => Number(a.session)-Number(b.session) || a.discipline.localeCompare(b.discipline))
  }

  function staffBusyInSession(staffId: string, day: string, session: string, currentAssignments = assignments) {
    if (!programme) return false
    if (arrivalStaffForDaySession(day, session).has(staffId)) return true
    if (Object.entries(waterSupportAssignments).some(([key, id]) => key.startsWith(`${day}::${session}::`) && id === staffId)) return true
    return programme.rows.filter((row) => row.day === day && row.session === session)
      .some((row) => activityCellsForRow(row).some((cell) => currentAssignments[cellKey(row.id, cell.group)] === staffId))
  }

  function setWaterSupport(day: string, session: string, discipline: 'canoe' | 'kayak', staffId: string) {
    const key = waterSupportKey(day, session, discipline)
    if (!staffId) {
      setWaterSupportAssignments((current) => {
        const next = { ...current }; delete next[key]
        localStorage.setItem(WATER_SUPPORT_KEY, JSON.stringify(next)); return next
      })
      return
    }
    const member = staff.find((item) => item.id === staffId)
    if (!member?.qualifications.includes(leadQualificationCode(discipline))) {
      setImportMessage(`${member?.name ?? 'That staff member'} is not signed off as ${discipline === 'canoe' ? 'Canoe Lead' : 'Kayak Lead'}.`)
      return
    }
    const assignedDisciplineGroups = disciplineAssignments(day, session, discipline)
    const ownGroup = assignedDisciplineGroups.find((item) => item.staffId === staffId)
    if (ownGroup) {
      setPendingWaterConfirmation({
        day, session, discipline, staffId,
        leadGroup: ownGroup.group,
        overseenGroups: assignedDisciplineGroups.filter((item) => item.group !== ownGroup.group).map((item) => item.group),
      })
      return
    }
    setWaterSupportAssignments((current) => {
      const next = { ...current, [key]: staffId }
      localStorage.setItem(WATER_SUPPORT_KEY, JSON.stringify(next)); return next
    })
  }

  async function confirmWaterLeadException(permissionFrom: 'Head of Centre' | 'Activities Manager') {
    const pending = pendingWaterConfirmation
    if (!pending) return
    const member = staff.find((item) => item.id === pending.staffId)
    if (!member) return
    const key = waterSupportKey(pending.day, pending.session, pending.discipline)
    setWaterSupportAssignments((current) => {
      const next = { ...current, [key]: pending.staffId }
      localStorage.setItem(WATER_SUPPORT_KEY, JSON.stringify(next)); return next
    })
    const { error } = await supabase.from('water_lead_logs').insert({
      programme_day: pending.day,
      session: pending.session,
      discipline: pending.discipline,
      lead_staff_id: member.id,
      lead_staff_name: member.name,
      lead_group: pending.leadGroup,
      overseen_groups: pending.overseenGroups,
      confirmed_by_name: displayName?.trim() || accountEmail,
      confirmed_by_email: accountEmail.trim().toLowerCase(),
      permission_from: permissionFrom,
    })
    if (error) setImportMessage(`Lead permission could not be logged: ${error.message}`)
    else setImportMessage(`${member.name} confirmed as ${pending.discipline === 'canoe' ? 'Canoe Lead' : 'Kayak Lead'} with ${permissionFrom} permission.`)
    setPendingWaterConfirmation(null)
    loadWaterLeadLogs()
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

    const workingIds = new Set(
      workingByDay[day] ?? staff.map((member) => member.id),
    )
    // Start from the current rota, but remove assignments that are now blocked
    // by the live Days Off & Sickness calendar. This prevents Auto-fill from
    // leaving SICK/OFF/HOL staff on sessions and applies AM/PM rules per session.
    const nextAssignments: StaffingAssignment = { ...assignments }
    const rowsForDay = programme.rows.filter((row) => row.day === day)
    for (const row of rowsForDay) {
      const blocked = unavailableStaffIdsForSession(day, row.session)
      for (const cell of activityCellsForRow(row)) {
        const key = cellKey(row.id, cell.group)
        const assignedStaffId = nextAssignments[key]
        if (assignedStaffId && blocked.has(assignedStaffId)) delete nextAssignments[key]
      }
    }

    const workload = new Map<string, number>()
    Object.entries(nextAssignments).forEach(([, staffId]) => {
      workload.set(staffId, (workload.get(staffId) ?? 0) + 1)
    })

    const dayRows = rowsForDay
    const clusters = staffingClustersForRows(dayRows).sort((a, b) => a.priority - b.priority || Number(a.session) - Number(b.session) || a.activityCode.localeCompare(b.activityCode))

    // Rebuild each activity cluster so shared activities only use the number of
    // instructors required by their saved staffing rule.
    for (const cluster of clusters) {
      cluster.cells.forEach(({ row, group }) => delete nextAssignments[cellKey(row.id, group)])

      const chosenStaff: StaffMember[] = []
      for (let slot = 0; slot < cluster.required; slot += 1) {
        const candidates = staff
          .filter((member) =>
            workingIds.has(member.id) &&
            !unavailableStaffIdsForSession(day, cluster.session).has(member.id) &&
            qualificationIsValid(member, cluster.activityCode) &&
            !arrivalStaffForDaySession(day, cluster.session).has(member.id) &&
            !chosenStaff.some((chosen) => chosen.id === member.id)
          )
          .filter((member) => !dayRows.some((otherRow) =>
            otherRow.session === cluster.session &&
            activityCellsForRow(otherRow).some((otherCell) => nextAssignments[cellKey(otherRow.id, otherCell.group)] === member.id)
          ))
          .sort((a, b) => {
            if (isWaterActivity(cluster.activityCode) && (cluster.session === '2' || cluster.session === '4')) {
              const aWater = hadWaterInPreviousPairedSession(a.id, day, cluster.session, nextAssignments)
              const bWater = hadWaterInPreviousPairedSession(b.id, day, cluster.session, nextAssignments)
              if (aWater !== bWater) return aWater ? -1 : 1
            }
            return rolePriority(resolvedRole(a)) - rolePriority(resolvedRole(b)) ||
              (workload.get(a.id) ?? 0) - (workload.get(b.id) ?? 0) ||
              a.name.localeCompare(b.name)
          })
        if (candidates[0]) chosenStaff.push(candidates[0])
      }

      if (!chosenStaff.length) continue
      // Balanced distribution: 10 groups / 3 staff becomes 4, 3 and 3.
      cluster.cells.forEach(({ row, group }, index) => {
        const staffIndex = Math.min(chosenStaff.length - 1, Math.floor(index * chosenStaff.length / cluster.cells.length))
        nextAssignments[cellKey(row.id, group)] = chosenStaff[staffIndex].id
      })
      chosenStaff.forEach((member) => workload.set(member.id, (workload.get(member.id) ?? 0) + 1))
    }

    const nextWaterSupport = { ...waterSupportAssignments }
    for (const need of waterSupportNeedsForDay(day)) {
      const supportKey = waterSupportKey(day, need.session, need.discipline)
      const currentId = nextWaterSupport[supportKey]
      const currentStillValid = currentId && staff.find((member) => member.id === currentId)?.qualifications.includes(leadQualificationCode(need.discipline)) && !unavailableStaffIdsForSession(day, need.session).has(currentId) && workingIds.has(currentId) && !arrivalStaffForDaySession(day, need.session).has(currentId) && (!staffBusyInSession(currentId, day, need.session, nextAssignments) || disciplineAssignments(day, need.session, need.discipline, nextAssignments).some((item) => item.staffId === currentId))
      if (currentStillValid) continue
      delete nextWaterSupport[supportKey]
      const qualificationCode = leadQualificationCode(need.discipline)
      const spareLead = staff.filter((member) =>
        member.qualifications.includes(qualificationCode) &&
        workingIds.has(member.id) &&
        !unavailableStaffIdsForSession(day, need.session).has(member.id) &&
        !staffBusyInSession(member.id, day, need.session, nextAssignments) &&
        !Object.entries(nextWaterSupport).some(([key, id]) => key.startsWith(`${day}::${need.session}::`) && id === member.id)
      ).sort((a,b) => rolePriority(resolvedRole(a))-rolePriority(resolvedRole(b)) || (workload.get(a.id)??0)-(workload.get(b.id)??0) || a.name.localeCompare(b.name))[0]
      if (spareLead) {
        nextWaterSupport[supportKey] = spareLead.id
        continue
      }
      const groupAssignments = disciplineAssignments(day, need.session, need.discipline, nextAssignments)
      const workingLead = groupAssignments
        .map((assignment) => ({ assignment, member: staff.find((item) => item.id === assignment.staffId) }))
        .find(({ member }) => member?.qualifications.includes(qualificationCode))
      if (workingLead?.member) {
        setPendingWaterConfirmation({
          day, session: need.session, discipline: need.discipline,
          staffId: workingLead.member.id,
          leadGroup: workingLead.assignment.group,
          overseenGroups: groupAssignments.filter((item) => item.group !== workingLead.assignment.group).map((item) => item.group),
        })
      } else {
        setImportMessage(`Water staffing warning: no qualified ${need.discipline === 'canoe' ? 'Canoe Lead' : 'Kayak Lead'} is available for ${day}, Session ${need.session}. Assign a qualified lead or remove one group.`)
      }
    }
    setWaterSupportAssignments(nextWaterSupport)
    localStorage.setItem(WATER_SUPPORT_KEY, JSON.stringify(nextWaterSupport))
    setAssignments(nextAssignments)
    localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(nextAssignments))
    setImportMessage(
      `Smart Auto-fill completed for ${day} using activity staffing rules, priorities and required canoe/kayak leads.`,
    )
  }

  function isDoubleBooked(
    staffId: string,
    targetRow: ProgrammeRow,
    targetGroup: number,
  ) {
    if (!programme) return false
    if (targetRow.session === '3' && arrivalStaffForDaySession(targetRow.day, targetRow.session).has(staffId)) return true
    const targetCell = activityCellsForRow(targetRow).find((cell) => cell.group === targetGroup)
    return programme.rows.some((row) =>
      row.day === targetRow.day && row.session === targetRow.session &&
      activityCellsForRow(row).some((cell) => {
        if (row.id === targetRow.id && cell.group === targetGroup) return false
        if (assignments[cellKey(row.id, cell.group)] !== staffId) return false
        const sameSharedCluster = targetCell && cell.activityCode === targetCell.activityCode &&
          (row.schoolLabel?.trim() || 'Centre programme') === (targetRow.schoolLabel?.trim() || 'Centre programme') &&
          staffingRuleForActivity(cell.activityCode).type !== 'per_group'
        return !sameSharedCluster
      }),
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
    const newActivity: Activity = {
      code,
      name,
      colour: '#dce8f5',
      equipmentQuantity: 0,
      capacity: 1,
      enabled: true,
      notes: '',
      staffingRuleType: 'per_group',
      staffingRuleValue: 1,
      staffingPriority: 2,
    }
    const next: Activity[] = [...activities, newActivity].sort((a, b) =>
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

  function updateActivityDefinition(code: string, patch: Partial<Activity>) {
    const next = activities.map((activity) => activity.code === code ? { ...activity, ...patch } : activity)
    setActivities(next)
    localStorage.setItem(ACTIVITIES_KEY, JSON.stringify(next))
  }

  function deleteActivity(code: string) {
    if (!window.confirm(`Remove ${code} from the activity list? Existing saved programmes may still contain this code. Disabling it is safer if you need to preserve old records.`)) return
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
    const sickIds = unavailableStaffIdsForSession(row.day, '3')
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
    const sickIds = unavailableStaffIdsForDay(day)
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
          !unavailableStaffIdsForSession(selectedStaffingCell.row.day, selectedStaffingCell.row.session).has(member.id) &&
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

  function programmeWeekKey(value: ProgrammeImport) {
    const source = `${value.title} ${value.sourceFileName}`
    const dateRange = source.match(/(\d{1,2})(?:st|nd|rd|th)?\s*[-–]\s*(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)(?:\s+(20\d{2}))?/i)
    if (dateRange) return `${dateRange[1]}-${dateRange[2]} ${dateRange[3]} ${dateRange[4] ?? new Date(value.importedAt).getFullYear()}`
    return value.title || value.sourceFileName || new Date(value.importedAt).toLocaleDateString('en-GB')
  }

  function migrateAssignments(previous: ProgrammeImport | null, next: ProgrammeImport, current: StaffingAssignment) {
    if (!previous) return current
    const migrated: StaffingAssignment = {}
    const oldRows = new Map(previous.rows.map((row) => [`${row.day}|${row.session}|${row.schoolLabel ?? ''}`, row]))
    for (const row of next.rows) {
      const old = oldRows.get(`${row.day}|${row.session}|${row.schoolLabel ?? ''}`) ?? previous.rows.find((item) => item.day === row.day && item.session === row.session)
      if (!old) continue
      for (const cell of row.cells) {
        const oldCell = old.cells.find((item) => item.group === cell.group)
        if (!oldCell || !cell.activityCode || cell.activityCode === 'Z') continue
        const assigned = current[cellKey(old.id, cell.group)]
        if (assigned) migrated[cellKey(row.id, cell.group)] = assigned
      }
    }
    return migrated
  }

  function archiveSnapshot(sourceProgramme = programme, force = false) {
    if (!sourceProgramme) return
    const weekKey = programmeWeekKey(sourceProgramme)
    if (!force && staffingArchives.some((item) => item.weekKey === weekKey)) return
    const archive: StaffingArchive = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      weekKey,
      title: sourceProgramme.title,
      sourceFileName: sourceProgramme.sourceFileName,
      archivedAt: new Date().toISOString(),
      archivedBy: displayName?.trim() || accountEmail || 'Manager',
      programme: structuredClone(sourceProgramme),
      assignments: structuredClone(assignments),
      workingByDay: structuredClone(workingByDay),
      sicknessByDay: structuredClone(sicknessByDay),
      daysOff: structuredClone(daysOff),
      staff: structuredClone(staff),
    }
    const next = [archive, ...staffingArchives].slice(0, 260)
    setStaffingArchives(next)
    localStorage.setItem(STAFFING_ARCHIVES_KEY, JSON.stringify(next))
    setImportMessage(`Staffing week ${weekKey} was archived permanently in Staffing Logs.`)
  }

  function xmlEscape(value: unknown) {
    return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
  }

  function sessionTime(session: string) {
    const times: Record<string, string> = {
      '1': '09:10–10:30',
      '2': '10:45–12:15',
      '3': '14:00–15:30',
      '4': '15:45–17:15',
      '5': '19:00–20:30',
    }
    return times[String(session)] ?? ''
  }

  function staffingArchiveMonth(archive: StaffingArchive) {
    const source = `${archive.weekKey} ${archive.title} ${archive.sourceFileName}`
    const months = ['january','february','march','april','may','june','july','august','september','october','november','december']
    const match = source.match(/(?:^|\s)(\d{1,2})(?:st|nd|rd|th)?(?:\s*[-–]\s*\d{1,2}(?:st|nd|rd|th)?)?\s+([A-Za-z]+)(?:\s+(20\d{2}))?/i)
    if (match) {
      const monthIndex = months.indexOf(match[2].toLowerCase())
      if (monthIndex >= 0) {
        const year = Number(match[3] ?? new Date(archive.programme.importedAt || archive.archivedAt).getFullYear())
        return `${year}-${String(monthIndex + 1).padStart(2, '0')}`
      }
    }
    const fallback = new Date(archive.programme.importedAt || archive.archivedAt)
    return `${fallback.getFullYear()}-${String(fallback.getMonth() + 1).padStart(2, '0')}`
  }

  function staffingMonthLabel(key: string) {
    const [year, month] = key.split('-').map(Number)
    return new Date(year, month - 1, 1).toLocaleDateString('en-GB', { month: 'long', year: 'numeric' })
  }

  function staffingColour(activityCode: string) {
    const text = `${activityCode} ${activityName(activityCode)}`.toLowerCase()
    if (/sail|canoe|kayak|paddle|sup|funboat|powerboat|safety boat|water/.test(text)) return 'water'
    if (/rope|climb|abseil|zip|crate|high|low/.test(text)) return 'ropes'
    return 'normal'
  }

  function dateForProgrammeDay(sourceProgramme: ProgrammeImport, day: string) {
    if (/^\d{4}-\d{2}-\d{2}$/.test(day)) return day
    if (sourceProgramme.startDate) { const rank = weekdayRank(day); const start = new Date(`${sourceProgramme.startDate}T12:00:00`); if (rank < 99 && !Number.isNaN(start.getTime())) { const startRank = ((start.getDay() + 6) % 7) + 1; const result = new Date(start); result.setDate(start.getDate() + ((rank - startRank + 7) % 7)); return result.toISOString().slice(0, 10) } }
    const source = `${sourceProgramme.title} ${sourceProgramme.sourceFileName}`
    const months = ['january','february','march','april','may','june','july','august','september','october','november','december']

    // Supports both "29th June - 5th July 2026" and "29th - 5th July 2026".
    const splitMonth = source.match(/(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s*[-–]\s*(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)(?:\s+(20\d{2}))?/i)
    const sameMonth = source.match(/(\d{1,2})(?:st|nd|rd|th)?\s*[-–]\s*(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)(?:\s+(20\d{2}))?/i)

    let firstDay: number
    let firstMonthName: string
    let explicitYear: string | undefined
    if (splitMonth) {
      firstDay = Number(splitMonth[1])
      firstMonthName = splitMonth[2]
      explicitYear = splitMonth[5]
    } else if (sameMonth) {
      firstDay = Number(sameMonth[1])
      firstMonthName = sameMonth[3]
      explicitYear = sameMonth[4]
    } else {
      return day
    }

    const month = months.indexOf(firstMonthName.toLowerCase())
    if (month < 0) return day
    const importedYear = new Date(sourceProgramme.importedAt).getFullYear()
    const year = Number(explicitYear ?? (Number.isFinite(importedYear) ? importedYear : new Date().getFullYear()))
    const first = new Date(year, month, firstDay)
    const wanted = ['SUN','MON','TUE','WED','THU','FRI','SAT'].findIndex((item) => day.toUpperCase().startsWith(item))
    if (wanted < 0) return day
    const date = new Date(first)
    date.setDate(first.getDate() + ((wanted - first.getDay() + 7) % 7))
    return dateKey(date)
  }

  function staffingStatus(member: StaffMember, day: string, sourceProgramme: ProgrammeImport, sourceDaysOff: StaffDayOff[], sourceWorking: Record<string,string[]>, _sourceSickness: Record<string,string[]>) {
    const isoDay = dateForProgrammeDay(sourceProgramme, day)
    const memberEmail = (member.email ?? '').trim().toLowerCase()
    const memberName = normaliseIdentity(member.name)
    const exact = sourceDaysOff.find((entry) => entry.day === isoDay && (
      entry.staff_id === member.id ||
      (memberEmail && entry.staff_email.trim().toLowerCase() === memberEmail) ||
      normaliseIdentity(entry.staff_name) === memberName
    ))
    if (exact?.status) return exact.status
    const working = sourceWorking[day]
    if (working && !working.includes(member.id)) return 'off' as DayOffStatus
    return null
  }

  async function createStaffingExcel(sourceProgramme: ProgrammeImport, sourceAssignments: StaffingAssignment, sourceStaff: StaffMember[], sourceDaysOff: StaffDayOff[], sourceWorking: Record<string,string[]>, sourceSickness: Record<string,string[]>, selectedDays: string[], fileName: string) {
    const ExcelJS = (window as Window & { ExcelJS?: any }).ExcelJS
    if (!ExcelJS) throw new Error('The Excel export library did not load. Refresh the page and try again.')

    const response = await fetch(`${import.meta.env.BASE_URL}staffing-template.xlsx`)
    if (!response.ok) throw new Error('The staffing Excel template could not be loaded.')

    // ExcelJS writes real cell styles (including fills) and produces an XLSX
    // package that Microsoft Excel opens without a repair warning.
    const templateWorkbook = new ExcelJS.Workbook()
    await templateWorkbook.xlsx.load(await response.arrayBuffer())
    const template = templateWorkbook.worksheets[0]
    if (!template) throw new Error('The staffing template worksheet is missing.')

    const normaliseDay = (value: string) => value.trim().toLowerCase().slice(0, 3)
    const programmeDays = Array.from(new Set(sourceProgramme.rows.map((row) => row.day)))
    const days = selectedDays
      .map((requested) => programmeDays.find((day) => normaliseDay(day) === normaliseDay(requested)))
      .filter((day): day is string => Boolean(day))
    if (!days.length) throw new Error('No matching programme days were found for this download.')

    const sessions = ['1', '2', '3', '4', '5']
    const sessionTimes = ['9:10-10:30', '10:45-12:15', '14:00-15:30', '15:45-17:15', '19:00-20:30']
    // The old template includes a redundant status column between the staff
    // name and Session 1. Export A, B and D:M so Session 1 starts immediately
    // after the name while preserving the existing timetable layout.
    const sourceColumns = [1, 2, ...Array.from({ length: 10 }, (_, index) => index + 4)]
    const blockWidth = sourceColumns.length
    const blockGap = 3
    const rowCount = 38
    const outputWorkbook = new ExcelJS.Workbook()
    outputWorkbook.creator = 'Norfolk Lakes'
    outputWorkbook.title = fileName.replace(/\.xlsx$/i, '')
    const output = outputWorkbook.addWorksheet('Staffing', {
      pageSetup: { ...template.pageSetup },
      properties: { ...template.properties },
      views: template.views ? JSON.parse(JSON.stringify(template.views)) : undefined,
    })

    const clone = <T,>(value: T): T => value == null ? value : JSON.parse(JSON.stringify(value)) as T

    // Copy only the single template table for every requested programme day.
    // This prevents hidden, duplicate and historic tables from entering exports.
    days.forEach((_, dayIndex) => {
      const destinationStart = dayIndex * (blockWidth + blockGap) + 1
      sourceColumns.forEach((sourceColumnNumber, outputColumnIndex) => {
        const sourceColumn = template.getColumn(sourceColumnNumber)
        const targetColumn = output.getColumn(destinationStart + outputColumnIndex)
        targetColumn.width = sourceColumn.width
        targetColumn.hidden = false
        targetColumn.outlineLevel = sourceColumn.outlineLevel
      })
      for (let row = 1; row <= rowCount; row += 1) {
        const sourceRow = template.getRow(row)
        const targetRow = output.getRow(row)
        if (dayIndex === 0) {
          targetRow.height = sourceRow.height
          targetRow.hidden = false
          targetRow.outlineLevel = sourceRow.outlineLevel
        }
        sourceColumns.forEach((sourceColumnNumber, outputColumnIndex) => {
          const sourceCell = sourceRow.getCell(sourceColumnNumber)
          const targetCell = targetRow.getCell(destinationStart + outputColumnIndex)
          targetCell.value = clone(sourceCell.value)
          targetCell.style = clone(sourceCell.style)
          if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt
          if (sourceCell.note) targetCell.note = clone(sourceCell.note)
        })
      }
    })

    const templateMerges: string[] = clone((template as any).model?.merges ?? [])
    const mappedColumn = (zeroBasedSourceColumn: number) => {
      const oneBased = zeroBasedSourceColumn + 1
      if (oneBased <= 2) return oneBased - 1
      if (oneBased === 3) return null
      if (oneBased <= 13) return oneBased - 2
      return null
    }
    days.forEach((_, dayIndex) => {
      const offset = dayIndex * (blockWidth + blockGap)
      templateMerges.forEach((merge) => {
        const range = XLSX.utils.decode_range(merge)
        if (range.s.r >= rowCount || range.e.r >= rowCount || range.s.c >= 13 || range.e.c >= 13) return
        let mappedStart = mappedColumn(range.s.c)
        let mappedEnd = mappedColumn(range.e.c)
        // A merge spanning the removed column remains a continuous merge after
        // the remaining columns are shifted left. A merge only in column C is dropped.
        if (mappedStart == null && mappedEnd == null) return
        if (mappedStart == null) mappedStart = mappedEnd
        if (mappedEnd == null) mappedEnd = mappedStart
        if (mappedStart == null || mappedEnd == null) return
        output.mergeCells(range.s.r + 1, mappedStart + 1 + offset, range.e.r + 1, mappedEnd + 1 + offset)
      })
    })

    const fills = {
      water: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9DC3E6' }, bgColor: { argb: 'FF9DC3E6' } },
      ropes: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4B6C2' }, bgColor: { argb: 'FFF4B6C2' } },
      hol: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' }, bgColor: { argb: 'FFC6E0B4' } },
      off: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' }, bgColor: { argb: 'FFFFE699' } },
      sick: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4CCCC' }, bgColor: { argb: 'FFF4CCCC' } },
    } as const

    const setCell = (row: number, column: number, value: string | number, fill: keyof typeof fills | null = null) => {
      const cell = output.getRow(row + 1).getCell(column + 1)
      cell.value = value
      // Keep borders, font, alignment and number formatting from the template.
      // Only replace the fill so activity/status colours are visible in Excel.
      cell.fill = fill ? clone(fills[fill]) : { type: 'pattern', pattern: 'none' }
    }

    const dutyFor = (member: StaffMember, day: string, session: string) => sourceProgramme.rows
      .filter((row) => row.day === day && String(row.session) === session)
      .flatMap((row) => row.cells
        .filter((cell) => sourceAssignments[cellKey(row.id, cell.group)] === member.id && cell.activityCode && cell.activityCode !== 'Z')
        .map((cell) => ({ code: cell.activityCode.toUpperCase(), group: cell.group })))

    days.forEach((day, dayIndex) => {
      const start = dayIndex * (blockWidth + blockGap)
      setCell(0, start, `DAILY STAFFING: ${day}`)
      sessionTimes.forEach((time, index) => setCell(1, start + 2 + index * 2, time))

      for (let row = 2; row < rowCount; row += 1) {
        for (let column = 0; column < blockWidth; column += 1) setCell(row, start + column, '')
      }

      const roleRank: Record<StaffRole, number> = { centreManager: 0, activityManager: 1, teamLeader: 2, staff: 3 }
      const orderedStaff = [...sourceStaff].sort((a, b) => roleRank[resolvedRole(a)] - roleRank[resolvedRole(b)] || a.name.localeCompare(b.name))
      orderedStaff.slice(0, 29).forEach((member, index) => {
        const row = index + 3
        const status = staffingStatus(member, day, sourceProgramme, sourceDaysOff, sourceWorking, sourceSickness)
        setCell(row, start, index + 1)
        setCell(row, start + 1, member.name)

        const statusLabel = status === 'hol' ? 'HOL' : status === 'sick' ? 'SICK' : status === 'am_off' ? 'AM OFF' : status === 'pm_off' ? 'PM OFF' : status === 'off' ? 'OFF' : ''
        const statusFill = status === 'hol' ? 'hol' : status === 'sick' ? 'sick' : status ? 'off' : null
        sessions.forEach((session, sessionIndex) => {
          const duties = dutyFor(member, day, session)
          const colours = duties.map((duty) => staffingColour(duty.code))
          const activityFill = colours.includes('water') ? 'water' : colours.includes('ropes') ? 'ropes' : null
          const sessionNumber = sessionIndex + 1
          const unavailable = Boolean(status) && (
            status === 'hol' || status === 'sick' || status === 'off' ||
            (status === 'am_off' && sessionNumber <= 2) ||
            (status === 'pm_off' && sessionNumber === 5)
          )

          // The live availability calendar wins over rota assignments. A blocked
          // session must contain only its status and must never export an activity.
          if (unavailable) {
            setCell(row, start + 2 + sessionIndex * 2, statusLabel, statusFill)
            setCell(row, start + 3 + sessionIndex * 2, '', statusFill)
          } else if (duties.length > 0) {
            setCell(row, start + 2 + sessionIndex * 2, duties.map((duty) => duty.code).join(' / '), activityFill)
            setCell(row, start + 3 + sessionIndex * 2, duties.map((duty) => `G${duty.group}`).join(', '), activityFill)
          } else {
            setCell(row, start + 2 + sessionIndex * 2, '')
            setCell(row, start + 3 + sessionIndex * 2, '')
          }
        })
      })
    })

    output.pageSetup = {
      ...clone(template.pageSetup),
      printArea: `A1:${XLSX.utils.encode_col((days.length - 1) * (blockWidth + blockGap) + blockWidth - 1)}${rowCount}`,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
    }
    output.headerFooter = clone(template.headerFooter)

    const bytes = await outputWorkbook.xlsx.writeBuffer()
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = URL.createObjectURL(blob)
    const anchor = document.createElement('a')
    anchor.href = url
    anchor.download = fileName
    document.body.appendChild(anchor)
    anchor.click()
    anchor.remove()
    window.setTimeout(() => URL.revokeObjectURL(url), 1000)
  }

  async function exportStaffingWeek() {
    if (!programme) return
    try {
      await createStaffingExcel(programme, assignments, staff, daysOff, workingByDay, sicknessByDay, programmeDays, `staffing-week-${programmeWeekKey(programme).replace(/[^a-z0-9]+/gi,'-')}.xlsx`)
    } catch (error) {
      setImportMessage(error instanceof Error ? error.message : 'The staffing workbook could not be downloaded.')
    }
  }

  async function exportArchivedStaffing(archive: StaffingArchive, day?: string) {
    const days = day ? [day] : Array.from(new Set(archive.programme.rows.map((row) => row.day)))
    try {
      await createStaffingExcel(archive.programme, archive.assignments, archive.staff, archive.daysOff, archive.workingByDay, archive.sicknessByDay, days, `staffing-${archive.weekKey.replace(/[^a-z0-9]+/gi,'-')}${day ? `-${day}` : ''}.xlsx`)
    } catch (error) {
      setImportMessage(error instanceof Error ? error.message : 'The archived staffing workbook could not be downloaded.')
    }
  }

  async function exportDailyStaffing(day: string) {
    if (!programme || !day) return
    try {
      await createStaffingExcel(programme, assignments, staff, daysOff, workingByDay, sicknessByDay, [day], `${day}-daily-staffing.xlsx`)
    } catch (error) {
      setImportMessage(error instanceof Error ? error.message : 'The daily staffing workbook could not be downloaded.')
    }
  }

  const filteredStaffingCells = populatedCells.filter(({ row, cell }) => {
    if (activeStaffingDay && row.day !== activeStaffingDay) return false
    const staffId = assignments[cellKey(row.id, cell.group)]
    const staffName = staff.find((member) => member.id === staffId)?.name ?? ''
    return `${row.day} ${row.session} group ${cell.group} ${cell.activityCode} ${activityName(cell.activityCode)} ${staffName}`
      .toLowerCase()
      .includes(query.toLowerCase())
  })

  const staffingCalendarCells = populatedCells.filter(({ row, cell }) => {
    const staffId = assignments[cellKey(row.id, cell.group)]
    const staffName = staff.find((member) => member.id === staffId)?.name ?? ''
    return `${row.day} ${row.session} group ${cell.group} ${cell.activityCode} ${activityName(cell.activityCode)} ${staffName}`
      .toLowerCase()
      .includes(query.toLowerCase())
  })
  const staffingCalendarSessions = Array.from(
    new Set(staffingCalendarCells.map(({ row }) => row.session)),
  ).sort((a, b) => Number(a) - Number(b))

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

  const hasProgramme = Boolean(programme && programme.rows.length > 0)
  const schoolsOnSite = hasProgramme ? new Set(programme?.rows.map(arrivalSchoolName).filter(Boolean) ?? []).size : 0
  const todayAvailabilityDay = new Date().toISOString().slice(0, 10)
  const todayWorkingIds = workingByDay[todayAvailabilityDay]
  const todayUnavailableIds = unavailableStaffIdsForDay(todayAvailabilityDay)
  const availableTodayCount = staff.filter((member) =>
    (!todayWorkingIds || todayWorkingIds.includes(member.id)) && !todayUnavailableIds.has(member.id),
  ).length

  // Live dashboard count based on each staff member's individual status.
  // Working: 08:30–20:30, AM OFF: 12:15–20:30, PM OFF: 08:30–17:15.
  // OFF, holiday and sick are never counted.
  const staffCurrentlyOnSiteCount = (() => {
    const today = dateKey(currentTime)
    const minutesNow = currentTime.getHours() * 60 + currentTime.getMinutes()
    const centreOpen = 8 * 60 + 30
    const amOffStart = 12 * 60 + 15
    const pmOffEnd = 17 * 60 + 15
    const centreClose = 20 * 60 + 30

    if (minutesNow < centreOpen || minutesNow > centreClose) return 0

    const workingIds = workingByDay[today]
    const sickIds = new Set(sicknessByDay[today] ?? [])

    return staff.filter((member) => {
      if (workingIds && !workingIds.includes(member.id)) return false
      if (sickIds.has(member.id)) return false

      const status = daysOff.find((entry) => entry.staff_id === member.id && entry.day === today)?.status
      if (status === 'off' || status === 'hol' || status === 'sick') return false
      if (status === 'am_off') return minutesNow >= amOffStart && minutesNow <= centreClose
      if (status === 'pm_off') return minutesNow >= centreOpen && minutesNow <= pmOffEnd
      return minutesNow >= centreOpen && minutesNow <= centreClose
    }).length
  })()
  const dailyShortages = programmeDays.map((day) => {
    const required = busiestSessionForDay(day)?.total ?? 0
    const available = (workingByDay[day] ?? staff.map((member) => member.id))
      .filter((id) => !unavailableStaffIdsForDay(day).has(id)).length
    return { day, required, available, shortfall: Math.max(0, required - available) }
  })
  const staffingShortages = new Set([
    ...dailyShortages.filter((item) => item.shortfall > 0).map((item) => item.day),
    ...programmeDays.filter((day) => qualificationShortagesForDay(day).length > 0),
  ]).size
  const selectedDayCapacityShortfall = dailyShortages.find((item) => item.day === activeStaffingDay)?.shortfall ?? 0
  const selectedDayQualificationShortages = qualificationShortagesForDay(activeStaffingDay)
  const programmeQualificationShortageDays = programmeDays.filter((day) => qualificationShortagesForDay(day).length > 0).length

  const plannedMySessions = useMemo<MySessionDuty[]>(() => {
    if (!programme || !myStaffMember) return []
    const duties: MySessionDuty[] = []
    programme.rows.forEach((row) => {
      activityCellsForRow(row).forEach((cell) => {
        if (assignments[cellKey(row.id, cell.group)] !== myStaffMember.id) return
        duties.push({
          id: `planned-${row.id}-${cell.group}`,
          programme_name: programme.title,
          day: row.day,
          session: row.session,
          activity_name: activityName(cell.activityCode),
          group_numbers: [cell.group],
          duty_type: 'activity',
          school_name: arrivalSchoolName(row) || null,
          building_name: null,
          party_leader_name: null,
        })
      })
    })
    arrivalRows.forEach((row) => {
      const details = arrivalAssignment(row)
      if (details.leaderId === myStaffMember.id) {
        duties.push({ id: `planned-leader-${row.id}`, programme_name: programme.title, day: row.day, session: row.session, activity_name: 'Party Leader', group_numbers: [], duty_type: 'arrival_leader', school_name: arrivalSchoolName(row) || null, building_name: accommodationSummary(details.flatIds) || null, party_leader_name: myStaffMember.name })
      }
      const groups = row.cells.filter((_, index) => details.guideIds[index] === myStaffMember.id).map((cell) => cell.group)
      if (groups.length) duties.push({ id: `planned-arrival-${row.id}`, programme_name: programme.title, day: row.day, session: row.session, activity_name: 'Accommodation', group_numbers: groups, duty_type: 'arrival_instructor', school_name: arrivalSchoolName(row) || null, building_name: accommodationSummary(details.flatIds) || null, party_leader_name: staff.find((member) => member.id === details.leaderId)?.name ?? null })
    })
    return duties
  }, [programme, myStaffMember, assignments, arrivalAssignments, staff])

  const effectiveMySessions = mySessions.length ? mySessions : plannedMySessions

  useEffect(() => {
    if (!effectiveMySessions.length) return
    if (!selectedMySessionsDay || !effectiveMySessions.some((duty) => duty.day === selectedMySessionsDay)) {
      setSelectedMySessionsDay(effectiveMySessions[0].day)
    }
  }, [effectiveMySessions, selectedMySessionsDay])

  return (
    <div className="app-shell">
      <input
        ref={fileInputRef}
        className="hidden-input"
        type="file"
        accept=".xlsx,.xls"
        onChange={importExcel}
      />
      <input
        ref={appProgrammeInputRef}
        className="hidden-input"
        type="file"
        accept=".xlsx"
        onChange={importAppProgramme}
      />

      <header className="topbar">
        <div>
          <p className="eyebrow">Norfolk Lakes</p>
          <div className="brand-lockup"><img src={`${import.meta.env.BASE_URL}manor-adventure-logo.png`} alt="Manor Adventure"/><div><div className="brand-title-row"><h1>Adventure Centre Manager</h1><span className="release-pill">v5.0.0</span></div><small>Norfolk Lakes</small></div></div>
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
          <button className="display-top" onClick={() => { const url = new URL(window.location.href); url.searchParams.set('display', 'staff-room'); window.open(url.toString(), '_blank', 'noopener,noreferrer') }}>
            <Monitor size={17} />
            Staff room display
          </button>
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

      <Nav page={page} setPage={setPage} accountRole={accountRole} />

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
                <strong>{staffCurrentlyOnSiteCount}</strong>
                <small>staff currently on site</small>
                <div className={!hasProgramme ? 'hero-alert' : staffingShortages ? 'hero-alert warning' : 'hero-alert ready'}>
                  {!hasProgramme ? 'No programme loaded' : staffingShortages ? `${staffingShortages} staffing gaps need attention` : 'Programme fully staffed'}
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
              <Stat icon={<UserRoundCheck />} value={staffCurrentlyOnSiteCount} label="Staff currently on site" />
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
              <div className="programme-upload-actions">
                <button className="primary" onClick={() => { setProgrammeBuilderScreen('library'); setPage('programmeBuilder') }}>
                  <CalendarRange size={18} /> Programme Library
                </button>
                <button className="secondary-action" onClick={() => fileInputRef.current?.click()}>
                  <Upload size={18} /> Upload Programme
                </button>
              </div>
              {programme && (
                <>
                  <div className="programme-details">
                    <strong>{programme.title || programme.sourceFileName}</strong>
                    <span>{friendlyProgrammeDateRange(programme.startDate ?? '', programme.endDate ?? '')} · Last updated {new Date(programme.importedAt).toLocaleString()}</span>
                  </div>
                  <div className="programme-toolbar-actions">
                    <button className="secondary-action" onClick={() => setPage('programmeBuilder')}><CalendarRange size={18}/>Edit programme</button>
                    <button className="primary" onClick={() => void downloadPublishedProgrammeExcel()}><FileSpreadsheet size={18}/>Download Excel</button>
                  </div>
                </>
              )}
            </div>

            {!programme ? (
              <>
                <section className="staffing-no-programme-banner">
                  <div>
                    <p className="eyebrow">Availability mode</p>
                    <h3>No programme loaded</h3>
                    <p>You can still view and edit this week's staff days off, holidays and sickness. Uploading a programme will restore the normal activity staffing tools.</p>
                  </div>
                  <button className="secondary-action" onClick={() => fileInputRef.current?.click()}><Upload size={18}/>Upload programme</button>
                </section>
                <div className="day-tabs staffing-day-tabs" role="tablist" aria-label="Availability day">
                  {availabilityWeekDays.map((day) => (
                    <button key={day} className={activeStaffingDay === day ? 'active' : ''} onClick={() => setSelectedStaffingDay(day)}>
                      {new Date(`${day}T12:00:00`).toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' })}
                    </button>
                  ))}
                </div>
                <section className="staffing-availability-panel">
                  <div className="staffing-availability-heading">
                    <div><p className="eyebrow">Staff availability</p><h3>Days Off &amp; Sickness</h3><p>{new Date(`${activeStaffingDay}T12:00:00`).toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</p></div>
                    <button className="secondary-action" onClick={() => setPage('holidays')}>Open full calendar</button>
                  </div>
                  <div className="staffing-availability-table">
                    <div className="availability-table-head">Staff member</div>
                    <div className="availability-table-head">Role</div>
                    <div className="availability-table-head">Status</div>
                    {sortedDaysOffStaff().map((member) => {
                      const existing = daysOff.find((entry) => memberIdForDayOff(entry) === member.id && entry.day === activeStaffingDay)
                      const value: DayOffStatus | 'working' = existing?.status ?? 'working'
                      return <Fragment key={`no-programme-availability-${member.id}`}>
                        <div className="availability-staff-name"><strong>{member.name}</strong></div>
                        <div><span className="role-pill">{roleLabel(resolvedRole(member))}</span></div>
                        <div className={`availability-status status-${value}`}>
                          <select value={value} disabled={!canManageHolidays && accountRole !== 'teamLeader'} onChange={(event) => void setSingleDayOff(member, activeStaffingDay, event.target.value as DayOffStatus | 'working')}>
                            <option value="working">Working</option>
                            {canManageHolidays && <option value="off">OFF</option>}
                            {canManageHolidays && <option value="hol">HOL</option>}
                            <option value="sick">SICK</option>
                            {canManageHolidays && <option value="am_off">AM OFF</option>}
                            {canManageHolidays && <option value="pm_off">PM OFF</option>}
                          </select>
                        </div>
                      </Fragment>
                    })}
                  </div>
                </section>
              </>
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

        {page === 'programmeArchive' && (
          <Panel title="Programme Files" onBack={() => setPage('admin')}>
            <section className="programme-archive-intro">
              <div><p className="eyebrow">Admin programme filing</p><h3>Search every programme by school, year and month</h3><p>Programmes are filed automatically when they are loaded into the app. Download a complete centre programme or an individual school programme.</p></div>
              <span className="release-pill">v2.10</span>
            </section>
            <div className="programme-archive-filters">
              <label className="archive-search"><Search size={18}/><input value={archiveSearch} onChange={(event) => setArchiveSearch(event.target.value)} placeholder="Search school or programme"/></label>
              <label>Year<select value={archiveYear} onChange={(event) => { setArchiveYear(event.target.value); setArchiveMonth('') }}><option value="">All years</option>{archiveYears.map((year) => <option key={year} value={year}>{year}</option>)}</select></label>
              <label>Month<select value={archiveMonth} onChange={(event) => setArchiveMonth(event.target.value)}><option value="">All months</option>{Array.from({ length: 12 }, (_, index) => String(index + 1).padStart(2, '0')).map((month) => <option key={month} value={month}>{new Date(2026, Number(month) - 1, 1).toLocaleDateString('en-GB', { month: 'long' })}</option>)}</select></label>
              {(archiveSearch || archiveYear || archiveMonth) && <button className="secondary-action" onClick={() => { setArchiveSearch(''); setArchiveYear(''); setArchiveMonth('') }}>Clear filters</button>}
            </div>
            {!filteredProgrammeArchive.length ? <div className="empty-state"><History size={34}/><h3>No programme files found</h3><p>Load or upload a programme and it will appear here automatically.</p></div> : <div className="programme-archive-list">{filteredProgrammeArchive.map((item) => <article className="programme-archive-card" key={item.archiveId}>
              <header><div><span>{friendlyProgrammeDateRange(item.startDate ?? '', item.endDate ?? '')}</span><h3>{item.title || item.sourceFileName}</h3><p>{item.schoolDetails?.length ?? 0} school{(item.schoolDetails?.length ?? 0) === 1 ? '' : 's'} · {item.groupNumbers.length} groups · Filed {new Date(item.archivedAt).toLocaleDateString('en-GB')}</p></div><div className="archive-main-actions"><button className="primary" onClick={() => void downloadPublishedProgrammeExcel(item)}><FileSpreadsheet size={17}/>Complete Excel</button><button className="secondary-action" onClick={() => printArchivedProgramme(item)}><Printer size={17}/>Complete PDF</button><button className="icon-button small" title="Delete archive" onClick={() => deleteArchivedProgramme(item.archiveId)}><Trash2 size={16}/></button></div></header>
              <div className="archive-school-list">{(item.schoolDetails ?? []).map((school) => <div className="archive-school-row" key={school.id || school.schoolName}><div><strong>{school.schoolName}</strong><span>{school.groupNumbers?.length ?? 0} groups · {friendlyProgrammeDateRange(school.arrivalDate || item.startDate || '', school.departureDate || item.endDate || '')}</span></div><div><button className="secondary-action" onClick={() => void downloadPublishedProgrammeExcel(schoolProgrammeFromArchive(item, school.schoolName))}><FileSpreadsheet size={16}/>Excel</button><button className="secondary-action" onClick={() => printArchivedProgramme(item, school.schoolName)}><Printer size={16}/>PDF</button></div></div>)}</div>
            </article>)}</div>}
          </Panel>
        )}

        {page === 'arrivals' && (
          <Panel title="Arrivals" onBack={() => setPage('dashboard')}>
            {!programme ? (
              <>
                <section className="staffing-no-programme-banner">
                  <div>
                    <p className="eyebrow">Calendar mode</p>
                    <h3>No programme loaded for these dates</h3>
                    <p>The staffing calendar remains available so days off, holidays and sickness can still be viewed and managed.</p>
                  </div>
                  <button className="secondary-action" onClick={() => fileInputRef.current?.click()}><Upload size={18}/>Upload programme</button>
                </section>
                <div className="day-tabs staffing-day-tabs" role="tablist" aria-label="Staffing calendar dates">
                  {availabilityWeekDays.map((day) => (
                    <button key={day} className={activeStaffingDay === day ? 'active' : ''} onClick={() => setSelectedStaffingDay(day)}>
                      {new Date(`${day}T12:00:00`).toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' })}
                    </button>
                  ))}
                </div>
                <section className="staffing-availability-panel">
                  <div className="staffing-availability-heading">
                    <div>
                      <p className="eyebrow">Daily staffing calendar</p>
                      <h3>{new Date(`${activeStaffingDay}T12:00:00`).toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</h3>
                      <p>No programme is loaded for this date. Staff availability can still be managed.</p>
                    </div>
                    <button className="primary" onClick={() => setPage('holidays')}><CalendarDays size={18}/>Open full calendar</button>
                  </div>
                  <div className="staffing-availability-summary" aria-label="Staff availability summary">
                    <article><span>Staff total</span><strong>{staff.length}</strong></article>
                    <article><span>Available</span><strong>{staff.length - unavailableStaffIdsForDay(activeStaffingDay).size}</strong></article>
                    <article><span>Unavailable</span><strong>{unavailableStaffIdsForDay(activeStaffingDay).size}</strong></article>
                  </div>
                </section>
              </>
            ) : (
              <>
                <section className="arrivals-module-intro">
                  <div>
                    <p className="eyebrow">Programme-driven arrivals</p>
                    <h3>Monday, Wednesday and Friday · Session 3</h3>
                    <p>School names are taken directly from the uploaded programme. Allocate each school to accommodation, choose its Party Leader and staff the groups here.</p>
                  </div>
                  <span className="release-pill">v3.01</span>
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
                        const unavailableToday = unavailableStaffIdsForDay(activeStaffingDay)
                        const usedElsewhere = arrivalStaffUsedByOtherSchools(row)

                        const leaderOptions = staff.filter((member) => availableToday.includes(member.id) && !unavailableToday.has(member.id) && ['staff', 'teamLeader'].includes(resolvedRole(member)) && !usedElsewhere.has(member.id) && !assignment.guideIds.includes(member.id))
                        const guideOptions = staff.filter((member) => availableToday.includes(member.id) && !unavailableToday.has(member.id) && member.id !== assignment.leaderId && !usedElsewhere.has(member.id) && (arrivalStaffGroup === 'all' || resolvedRole(member) === arrivalStaffGroup))

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
            {programme && <div className="staffing-date-context"><article><span>Programme</span><strong>{programme.title}</strong><small>{programmeDays.length ? friendlyProgrammeDateRange(dateForProgrammeDay(programme, programmeDays[0]), dateForProgrammeDay(programme, programmeDays[programmeDays.length - 1])) : programme.title}</small></article><article className="today-card"><span>Selected programme date</span><strong>{new Date(`${dateForProgrammeDay(programme, activeStaffingDay)}T12:00:00`).toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</strong></article></div>}
            {!programme ? (
              <EmptyProgramme onUpload={() => fileInputRef.current?.click()} />
            ) : (
              <>
                <div className="staffing-view-toolbar">
                  <div className="staffing-view-switch" role="group" aria-label="Staffing view">
                    <button className={staffingView === 'activity' ? 'active' : ''} onClick={() => setStaffingView('activity')}>Activity View</button>
                    <button className={staffingView === 'calendar' ? 'active' : ''} onClick={() => setStaffingView('calendar')}>Calendar View</button>
                    <button onClick={() => setPage('holidays')}>Days Off &amp; Sickness</button>
                  </div>
                  <div className="staffing-zoom-controls" role="group" aria-label="Staffing zoom">
                    <button aria-label="Zoom out" disabled={staffingZoom <= 70} onClick={() => setStaffingZoom((current) => { const next = Math.max(70, current - 10); localStorage.setItem('acm-staffing-zoom', String(next)); return next })}>−</button>
                    <strong>{staffingZoom}%</strong>
                    <button aria-label="Zoom in" disabled={staffingZoom >= 150} onClick={() => setStaffingZoom((current) => { const next = Math.min(150, current + 10); localStorage.setItem('acm-staffing-zoom', String(next)); return next })}>+</button>
                    <button className="zoom-reset" onClick={() => { setStaffingZoom(100); localStorage.setItem('acm-staffing-zoom', '100') }}>Reset</button>
                  </div>
                </div>
                <div className="staffing-controls">
                  <div className="staffing-day-row">
                  <div className="day-tabs staffing-day-tabs" role="tablist" aria-label="Staffing day">
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
                  </div>

                  <div className="staffing-availability-summary" aria-label="Staffing availability summary">
                    <article><span>Staff total</span><strong>{staff.length}</strong></article>
                    <article><span>In today</span><strong>{staff.length - unavailableStaffIdsForDay(activeStaffingDay).size}</strong></article>
                    <article><span>OFF / HOL</span><strong>{(() => {
                      const availabilityDay = programme ? dateForProgrammeDay(programme, activeStaffingDay) : activeStaffingDay
                      return new Set(daysOff
                        .filter((entry) => entry.day === availabilityDay && entry.status !== 'sick')
                        .map((entry) => memberIdForDayOff(entry) ?? entry.staff_id)).size
                    })()}</strong></article>
                    <article className="sick-summary"><span>Sick</span><strong>{(() => {
                      const availabilityDay = programme ? dateForProgrammeDay(programme, activeStaffingDay) : activeStaffingDay
                      return new Set(
                        daysOff
                          .filter((entry) => entry.day === availabilityDay && entry.status === 'sick')
                          .map((entry) => memberIdForDayOff(entry) ?? entry.staff_id),
                      ).size
                    })()}</strong></article>
                    <button className="clear-staffing-button" onClick={() => clearDayStaffing(activeStaffingDay)}><X size={17}/>Clear day</button>
                    <button className="auto-fill-button" onClick={() => autoFillStaffing(activeStaffingDay)}><WandSparkles size={17}/>Auto-fill staff</button>
                  </div>
                  <div className="staffing-actions staffing-download-row">
                    <button className="print-button" onClick={() => void exportDailyStaffing(activeStaffingDay)}><FileSpreadsheet size={17}/>Download day</button>
                    <button className="print-button" onClick={() => void exportStaffingWeek()}><FileSpreadsheet size={17}/>Download full week</button>
                    {canViewLogs && <button className="secondary-action" onClick={() => archiveSnapshot(programme, true)}><History size={17}/>Archive week</button>}
                  </div>
                </div>

                {selectedDayCapacityShortfall > 0 && (
                  <section className="staffing-shortage-alert" role="alert">
                    <CircleAlert size={24} />
                    <div>
                      <strong>Not enough staff in today</strong>
                      <p>
                        {`${activeStaffingDay} cannot be fully staffed. You need ${selectedDayCapacityShortfall} more instructor${selectedDayCapacityShortfall === 1 ? '' : 's'} to cover every activity. ${selectedDayBusiest?.total ?? 0} required, ${availableTodayCount} available.`}
                      </p>
                    </div>
                  </section>
                )}

                {selectedDayQualificationShortages.length > 0 && (
                  <section className="qualification-shortage-alert" role="alert">
                    <CircleAlert size={24} />
                    <div>
                      <strong>Not enough correctly signed-off staff</strong>
                      <p>{activeStaffingDay} has activities that cannot be safely covered by the available sign-offs.</p>
                      <ul>
                        {selectedDayQualificationShortages.map((item) => (
                          <li key={`${item.session}-${item.activityCode}`}>
                            <b>Session {item.session} · {item.activityName}:</b> {item.required} needed, {item.covered} qualified available — {item.shortfall} more required.
                          </li>
                        ))}
                      </ul>
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

                {waterSupportNeedsForDay(activeStaffingDay).length > 0 && (
                  <section className="water-support-panel">
                    <div><p className="eyebrow">Water ratios</p><h3>Water Support</h3><p>Extra leads are required when two or more groups run the same canoe or kayak activity in one session.</p></div>
                    <div className="water-support-list">
                      {waterSupportNeedsForDay(activeStaffingDay).map((need) => {
                        const key = waterSupportKey(activeStaffingDay, need.session, need.discipline)
                        const assignedId = waterSupportAssignments[key] ?? ''
                        const unavailable = unavailableStaffIdsForDay(activeStaffingDay)
                        const working = new Set(workingByDay[activeStaffingDay] ?? staff.map((member) => member.id))
                        return <article key={key}>
                          <div><strong>Session {need.session} · {need.discipline === 'canoe' ? 'Canoe Lead' : 'Kayak Lead'}</strong><span>{need.groups} groups running</span></div>
                          <select value={assignedId} onChange={(event) => setWaterSupport(activeStaffingDay, need.session, need.discipline, event.target.value)}>
                            <option value="">Select lead</option>
                            {qualifiedWaterLeads(need.discipline).filter((member) => working.has(member.id) && !unavailable.has(member.id) && (!staffBusyInSession(member.id, activeStaffingDay, need.session) || member.id === assignedId || disciplineAssignments(activeStaffingDay, need.session, need.discipline).some((item) => item.staffId === member.id))).map((member) => <option key={member.id} value={member.id}>{member.name}{disciplineAssignments(activeStaffingDay, need.session, need.discipline).some((item) => item.staffId === member.id) ? ' · running a group' : ''}</option>)}
                          </select>
                        </article>
                      })}
                    </div>
                  </section>
                )}

                {staffingView === 'availability' && (
                  <section className="staffing-availability-panel">
                    <div className="staffing-availability-heading">
                      <div><p className="eyebrow">Single source of truth</p><h3>Days Off &amp; Sickness — {activeStaffingDay}</h3><p>Changes here immediately control Auto-fill, manual assignments, shortage warnings and Excel downloads.</p></div>
                    </div>
                    <div className="staffing-availability-table">
                      <div className="availability-table-head">Staff member</div>
                      <div className="availability-table-head">Role</div>
                      <div className="availability-table-head">Status</div>
                      {sortedDaysOffStaff().map((member) => {
                        const availabilityDay = programme ? dateForProgrammeDay(programme, activeStaffingDay) : activeStaffingDay
                        const existing = daysOff.find((entry) => memberIdForDayOff(entry) === member.id && entry.day === availabilityDay)
                        const legacySick = [...(sicknessByDay[activeStaffingDay] ?? []), ...(sicknessByDay[availabilityDay] ?? [])].includes(member.id)
                        const value: DayOffStatus | 'working' = existing?.status ?? (legacySick ? 'sick' : 'working')
                        return <Fragment key={`staffing-availability-${member.id}`}>
                          <div className="availability-staff-name"><strong>{member.name}</strong></div>
                          <div><span className="role-pill">{roleLabel(resolvedRole(member))}</span></div>
                          <div className={`availability-status status-${value}`}>
                            <select value={value} disabled={!canManageHolidays && accountRole !== 'teamLeader'} onChange={(event) => void setSingleDayOff(member, activeStaffingDay, event.target.value as DayOffStatus | 'working')}>
                              <option value="working">Working</option>
                              {canManageHolidays && <option value="off">OFF</option>}
                              {canManageHolidays && <option value="hol">HOL</option>}
                              <option value="sick">SICK</option>
                              {canManageHolidays && <option value="am_off">AM OFF</option>}
                              {canManageHolidays && <option value="pm_off">PM OFF</option>}
                            </select>
                          </div>
                        </Fragment>
                      })}
                    </div>
                  </section>
                )}

                {staffingView !== 'availability' && (
                <div className="staffing-zoom-stage" style={{ zoom: staffingZoom / 100 }}>
                {staffingView === 'activity' ? (
                  <div className="staffing-grid">
                    {filteredStaffingCells.map(({ row, cell }) => {
                      const key = cellKey(row.id, cell.group)
                      const assignedStaff = staff.find(
                        (member) => member.id === assignments[key],
                      )
                      const qualificationMissing = Boolean(assignedStaff && !qualificationIsValid(assignedStaff, cell.activityCode))
                      return (
                        <article
                          className={`staffing-card ${qualificationMissing ? 'qualification-missing' : assignedStaff ? 'ready' : 'needs'}`}
                          key={key}
                        >
                          <div className="staffing-card-top">
                            <span>
                              {row.day} · Session {row.session} · {sessionTime(row.session)}
                            </span>
                            <span>
                              {qualificationMissing ? 'Qualification missing' : assignedStaff ? 'Ready' : 'Needs instructor'}
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
                              {qualificationMissing && <small className="qualification-missing-text">Not signed off for {activityName(cell.activityCode)}</small>}
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
                ) : (
                  <section className="staffing-calendar-wrap" aria-label="Staffing calendar view">
                    <div
                      className="staffing-calendar"
                      style={{ gridTemplateColumns: `92px repeat(${programmeDays.length}, minmax(220px, 1fr))` }}
                    >
                      <div className="staffing-calendar-corner">Session</div>
                      {programmeDays.map((day) => (
                        <button
                          key={`calendar-day-${day}`}
                          className={`staffing-calendar-day ${day === activeStaffingDay ? 'active' : ''}`}
                          onClick={() => setSelectedStaffingDay(day)}
                        >
                          {day}
                        </button>
                      ))}
                      {staffingCalendarSessions.map((session) => (
                        <Fragment key={`calendar-session-${session}`}>
                          <div className="staffing-calendar-session"><strong>Session {session}</strong><span>{sessionTime(session)}</span></div>
                          {programmeDays.map((day) => {
                            const items = staffingCalendarCells.filter(
                              ({ row }) => row.day === day && row.session === session,
                            )
                            return (
                              <div className="staffing-calendar-cell" key={`${day}-${session}`}>
                                {items.length === 0 ? (
                                  <span className="staffing-calendar-empty">No activities</span>
                                ) : items.map(({ row, cell }) => {
                                  const key = cellKey(row.id, cell.group)
                                  const assignedStaff = staff.find((member) => member.id === assignments[key])
                                  const qualificationMissing = Boolean(
                                    assignedStaff && !qualificationIsValid(assignedStaff, cell.activityCode),
                                  )
                                  return (
                                    <button
                                      key={key}
                                      className={`staffing-calendar-card ${qualificationMissing ? 'qualification-missing' : assignedStaff ? 'ready' : 'needs'}`}
                                      onClick={() => setSelectedStaffingCell({ row, group: cell.group })}
                                    >
                                      <strong>{activityName(cell.activityCode)}</strong>
                                      <span>G{cell.group}</span>
                                      <small>{assignedStaff?.name ?? 'Needs instructor'}</small>
                                      {qualificationMissing && <em>Not signed off</em>}
                                    </button>
                                  )
                                })}
                              </div>
                            )
                          })}
                        </Fragment>
                      ))}
                    </div>
                  </section>
                )}
                </div>
                )}
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

        {page === 'programmeBuilder' && canManageStaff && (
          <Panel title="Programme Builder" onBack={() => programmeBuilderScreen === 'editor' ? setProgrammeBuilderScreen('library') : setPage('admin')}>
            {programmeBuilderScreen === 'library' ? (
            <section className="programme-library">
              <div className="programme-library-hero"><div><p className="eyebrow">Saved programme library</p><h2>Programmes</h2><p>Build a new programme, reopen an old one, add another school, or load it into Staffing and Arrivals.</p></div><button className="primary" onClick={createNewProgramme}><Plus size={18}/>Build new programme</button></div>
              <div className="programme-library-filters"><label><Search size={17}/><input value={programmeSearch} onChange={(event) => setProgrammeSearch(event.target.value)} placeholder="Search programme or school"/></label><label>Month<input type="month" value={programmeMonth} onChange={(event) => setProgrammeMonth(event.target.value)}/></label></div>
              {filteredSavedProgrammes.length ? <div className="programme-library-grid">{filteredSavedProgrammes.map((saved) => <article className="programme-library-card" key={saved.id}><div><span>{friendlyProgrammeDateRange(saved.startDate, saved.endDate)}</span><h3>{saved.title}</h3><p>{saved.draft.schools.length} school{saved.draft.schools.length === 1 ? '' : 's'} · Updated {new Date(saved.updatedAt).toLocaleDateString('en-GB')}</p><small>{saved.draft.schools.map((school) => school.name || school.programmeName).filter(Boolean).join(' · ') || 'No schools named yet'}</small></div><div className="programme-library-actions"><button className="secondary-action" onClick={() => openSavedProgramme(saved)}>Open and edit</button><button className="primary" onClick={() => loadSavedProgrammeIntoApp(saved)}>Load into app</button><button className="icon-button small" title="Delete saved programme" onClick={() => deleteSavedProgramme(saved.id)}><Trash2 size={16}/></button></div></article>)}</div> : <div className="empty-state"><CalendarRange size={34}/><h3>No saved programmes found</h3><p>Build your first programme or change the search and month filters.</p></div>}
            </section>
            ) : (
              <>
            <div className="builder-topbar"><button className="secondary-action" onClick={() => setProgrammeBuilderScreen('library')}><ChevronLeft size={17}/>Programme library</button>
              <div className="builder-mode-switch"><button className={programmeBuilderMode === 'design' ? 'active' : ''} onClick={() => setProgrammeBuilderMode('design')}>Design Programme</button><button className={programmeBuilderMode === 'upload' ? 'active' : ''} onClick={() => { setProgrammeBuilderMode('upload'); appProgrammeInputRef.current?.click() }}>Upload Programme</button></div>
              <div className="staffing-view-switch" role="group" aria-label="Programme builder view">
                <button className={programmeBuilderView === 'build' ? 'active' : ''} onClick={() => setProgrammeBuilderView('build')}>Build View</button>
                <button className={programmeBuilderView === 'preview' ? 'active' : ''} onClick={() => setProgrammeBuilderView('preview')}>Preview Calendar</button>
              </div>
              <div className="builder-actions">
                <button className="secondary-action" onClick={saveProgrammeToLibrary}><FileSpreadsheet size={18}/>Save programme</button>
                <button className="primary" onClick={publishProgrammeBuilder}><CheckCircle2 size={18}/>Publish programme</button>
              </div>
            </div>
            {programmeBuilderMessage && <p className="builder-message">{programmeBuilderMessage}</p>}

            {programmeBuilderView === 'build' ? (
              <>
                <section className="builder-settings-card">
                  <div className="builder-field-grid">
                    <label>Programme week name<input value={programmeBuilder.name} onChange={(event) => updateProgrammeBuilder({ name: event.target.value })} placeholder="e.g. 13–17 July 2026"/></label>
                    <label>Week start<input type="date" value={programmeBuilder.startDate} onChange={(event) => {
                      const startDate = event.target.value
                      const schools = programmeBuilder.schools.map((school) => school.arrivalDate === programmeBuilder.startDate ? { ...school, arrivalDate: startDate } : school)
                      updateProgrammeBuilder({ startDate, schools })
                    }}/></label>
                    <label>Week end<input type="date" min={programmeBuilder.startDate} value={programmeBuilder.endDate} onChange={(event) => {
                      const endDate = event.target.value
                      const schools = programmeBuilder.schools.map((school) => school.departureDate === programmeBuilder.endDate ? { ...school, departureDate: endDate } : school)
                      updateProgrammeBuilder({ endDate, schools })
                    }}/></label>
                  </div>
                </section>

                <section className="builder-section">
                  <div className="builder-section-heading"><div><p className="eyebrow">Schools and groups</p><h3>Who is attending?</h3></div><div className="builder-heading-actions"><button className="secondary-action" onClick={updateWholeProgramme}><WandSparkles size={17}/>Auto Fill Whole Programme</button><button className="secondary-action" onClick={addBuilderSchool}><Plus size={17}/>Add school</button></div></div>
                  <div className="builder-school-grid">{programmeBuilder.schools.map((school, index) => <article className="builder-school-card" key={school.id}>
                    <div className="builder-section-heading"><strong>School {index + 1}</strong><label className="builder-lock"><input type="checkbox" checked={school.locked} onChange={(event) => updateBuilderSchool(school.id, { locked: event.target.checked })}/>Lock programme</label></div>
                    <label>School name<input value={school.name} onChange={(event) => updateBuilderSchool(school.id, { name: event.target.value })} placeholder="School name"/></label>
                    <label>Programme name<input value={school.programmeName} onChange={(event) => updateBuilderSchool(school.id, { programmeName: event.target.value })} placeholder="e.g. Oakwood School – July"/></label>
                    <div className="builder-form-grid"><label>Purchase type<select value={school.purchaseType} onChange={(event) => updateBuilderSchool(school.id, { purchaseType: event.target.value as ProgrammePurchaseType })}><option value="normal">Normal Package</option><option value="bargain">Bargain Special</option></select></label><label>Arrival date<input type="date" min={programmeBuilder.startDate} max={programmeBuilder.endDate} value={school.arrivalDate} onChange={(event) => updateBuilderSchool(school.id, { arrivalDate: event.target.value })}/></label><label>Departure date<input type="date" min={school.arrivalDate || programmeBuilder.startDate} max={programmeBuilder.endDate} value={school.departureDate} onChange={(event) => updateBuilderSchool(school.id, { departureDate: event.target.value })}/></label></div>
                    <label>Programme notes<textarea rows={3} value={school.notes} onChange={(event) => updateBuilderSchool(school.id, { notes: event.target.value })} placeholder="Notes for this school only…"/></label>
                    <label>Number of groups<input type="number" min="1" max="30" value={school.groups} onChange={(event) => updateBuilderSchool(school.id, { groups: Math.max(1, Number(event.target.value) || 1) })}/></label>
                    <div><strong>First-choice activities</strong><p className="builder-help">Select the activities requested by the school. There is no separate first-choice field.</p><div className="builder-activity-chips">{activities.filter((activity) => activity.code !== 'Z').map((activity) => { const active = school.requestedActivities.includes(activity.code); return <button type="button" key={activity.code} className={active ? 'chip active' : 'chip'} onClick={() => updateBuilderSchool(school.id, { requestedActivities: active ? school.requestedActivities.filter((code) => code !== activity.code) : [...school.requestedActivities, activity.code] })}>{activity.code}<small>{activity.name}</small></button> })}</div></div>
                    <div className="builder-form-grid"><label>Backup option 1<select value={school.backupOption1} onChange={(event) => updateBuilderSchool(school.id, { backupOption1: event.target.value })}><option value="">No backup selected</option>{activities.filter((activity) => activity.code !== 'Z').map((activity) => <option key={activity.code} value={activity.code}>{activity.code} – {activity.name}</option>)}</select></label><label>Backup option 2<select value={school.backupOption2} onChange={(event) => updateBuilderSchool(school.id, { backupOption2: event.target.value })}><option value="">No backup selected</option>{activities.filter((activity) => activity.code !== 'Z').map((activity) => <option key={activity.code} value={activity.code}>{activity.code} – {activity.name}</option>)}</select></label></div>
                    <div className="builder-school-actions"><button className="secondary-action" onClick={() => autoFillProgrammeBuilder(school.id)}><WandSparkles size={17}/>Auto Fill School</button><button className="secondary-action" onClick={() => printSchoolProgramme(school.id)}><Printer size={17}/>School PDF</button><button className="secondary-action" onClick={() => void downloadSchoolProgrammeExcel(school.id)}><FileSpreadsheet size={17}/>School Excel</button></div>
                    {programmeBuilder.schools.length > 1 && <button className="icon-button small" title="Remove school" onClick={() => removeBuilderSchool(school.id)}><Trash2 size={16}/></button>}
                  </article>)}</div>
                </section>

                {programmeBuilder.schools.some((school) => school.purchaseType === 'bargain') && <section className="builder-section bargain-rules"><div className="builder-section-heading"><div><p className="eyebrow">Bargain Special rules</p><h3>Package limits</h3></div></div><label className="builder-limit-field">Maximum sessions per group<input type="number" min="1" max="35" value={programmeBuilder.bargainSessionLimit} onChange={(event) => updateProgrammeBuilder({ bargainSessionLimit: Math.max(1, Number(event.target.value) || 1) })}/></label><p>Select the activities included in this package. These can be updated when you provide the Bargain Special reference sheet.</p><div className="builder-activity-chips">{activities.map((activity) => { const active = programmeBuilder.bargainAllowedActivities.includes(activity.code); return <button key={activity.code} className={active ? 'chip active' : 'chip'} onClick={() => updateProgrammeBuilder({ bargainAllowedActivities: active ? programmeBuilder.bargainAllowedActivities.filter((code) => code !== activity.code) : [...programmeBuilder.bargainAllowedActivities, activity.code] })} title={activity.name}>{activity.code}<small>{activity.name}</small></button> })}</div></section>}

                <section className="builder-section">
                  <div className="builder-section-heading"><div><p className="eyebrow">Programme grid</p><h3>Assign activities</h3></div><span>{builderGroups.length} groups · {builderDays.length} days</span></div>
                  <div className="builder-grid-wrap"><table className="builder-grid"><thead><tr><th>Day</th><th>Session</th>{builderGroups.map(({ group, school }) => <th key={group}>G{group}<small>{school.name || 'School'}</small></th>)}</tr></thead><tbody>{builderDays.flatMap((dayInfo) => BUILDER_SESSIONS.map((session) => <tr key={`${dayInfo.day}-${session}`}><th>{dayInfo.label}</th><th>S{session}</th>{builderGroups.map(({ group, school }) => { const state = builderSchoolSessionState(school, dayInfo.date, session); const value = programmeBuilder.assignments[builderAssignmentKey(dayInfo.day, session, group)] ?? ''; const options = school.purchaseType === 'bargain' ? activities.filter((activity) => programmeBuilder.bargainAllowedActivities.includes(activity.code)) : activities; return <td key={group} className={`builder-state-${state}`} onDragOver={state === 'activity' ? (event) => event.preventDefault() : undefined} onDrop={state === 'activity' ? () => dropBuilderActivity(dayInfo.day, session, group, school.id) : undefined}>{state === 'arrival' ? <strong>{school.name || 'School'} – Arrival</strong> : state === 'departed' ? <span>Departed</span> : state === 'offsite' ? <span>Not on site</span> : <div className="builder-draggable-cell" draggable={Boolean(value)} onDragStart={() => value && setDraggedBuilderActivity({ key: builderAssignmentKey(dayInfo.day, session, group), code: value, schoolId: school.id })}><select className={programmeBuilder.manualLocks[builderAssignmentKey(dayInfo.day, session, group)] ? 'manual-locked' : ''} title={programmeBuilder.manualLocks[builderAssignmentKey(dayInfo.day, session, group)] ? 'Manual change locked' : 'Automatic activity'} value={value} onChange={(event) => setBuilderActivity(dayInfo.day, session, group, event.target.value)}><option value="">—</option>{options.filter((activity) => school.requestedActivities.includes(activity.code) || activity.code === value).map((activity) => <option key={activity.code} value={activity.code}>{activity.code} – {activity.name}</option>)}</select></div>}</td> })}</tr>))}</tbody></table></div>

                  <div className="builder-update-actions"><div><strong>Manual changes are protected</strong><p>Any activity you change is locked. Update Programme rearranges only the remaining sessions.</p></div><div><button className="secondary-action" onClick={resetProgrammeLocks}>Reset Locks</button><button className="primary" onClick={updateWholeProgramme}><WandSparkles size={17}/>Update Programme</button></div></div>
                </section>
              </>
            ) : (
              <section className="builder-preview">
                <header><div><p className="eyebrow">Multi-school programme</p><h3>{programmeBuilder.name || 'Untitled programme'}</h3><span>{friendlyProgrammeDateRange(programmeBuilder.startDate, programmeBuilder.endDate)} · {builderGroups.length} groups</span></div><button className="secondary-action" onClick={() => window.print()}><Printer size={18}/>Print preview</button></header>
                {builderValidation().length > 0 && <div className="builder-validation"><CircleAlert size={20}/><div><strong>Check before publishing</strong>{builderValidation().map((issue) => <p key={issue}>{issue}</p>)}</div></div>}
                <div className="builder-preview-days">{builderDays.map((dayInfo) => <article className="builder-preview-day" key={dayInfo.day}><h4>{dayInfo.label}</h4>{BUILDER_SESSIONS.map((session) => <section key={session}><strong>Session {session}</strong><div>{builderGroups.map(({ group, school }) => { const state = builderSchoolSessionState(school, dayInfo.date, session); const code = programmeBuilder.assignments[builderAssignmentKey(dayInfo.day, session, group)] ?? ''; const label = state === 'arrival' ? `${school.name || 'School'} – Arrival` : state === 'departed' ? 'Departed' : state === 'offsite' ? 'Not on site' : code ? `${code} — ${activityNameFromList(activities, code)}` : 'Not assigned'; return <button key={group} className={state === 'activity' && code ? 'filled' : 'empty'} onClick={() => setProgrammeBuilderView('build')}><span>G{group} · {school.name || 'School'}</span><b>{label}</b></button> })}</div></section>)}</article>)}</div>
              </section>
            )}

              </>
            )}
          </Panel>
        )}

        {page === 'admin' && (
          <Panel title="Admin" onBack={isAdminAccount ? undefined : () => setPage('dashboard')}>
            <section className="admin-choice-grid">
              <button className="admin-choice-card programme-builder-card" onClick={() => setPage('programmeArchive')}>
                <Archive size={34} />
                <div><h3>Programme Files</h3><p>Search programmes by school, year and month, then download school or complete centre files.</p></div>
              </button>
              {!isAdminAccount && <section className="display-manager-card"><div><Monitor size={34}/><div><h3>Display Manager</h3><p>Open or copy the live read-only links for screens around the centre.</p></div></div><div className="display-link-list">{([['Staff room','staff-room'],['Manager','manager'],['Programme','programme']] as const).map(([label,mode]) => { const url = new URL(window.location.href); url.searchParams.set('display',mode); return <article key={mode}><strong>{label}</strong><code>{url.toString()}</code><button onClick={() => window.open(url.toString(),'_blank','noopener,noreferrer')}>Open</button><button onClick={() => { void navigator.clipboard.writeText(url.toString()); setImportMessage(`${label} display link copied.`) }}>Copy link</button></article> })}</div></section>}
              {canManageStaff && <button className="admin-choice-card programme-builder-card" onClick={() => setPage('programmeBuilder')}>
                <CalendarRange size={34} />
                <div><h3>Programme Builder</h3><p>Design Bargain Special or Normal Purchase programmes, preview them and publish them.</p></div>
              </button>}
              {canManageStaff && <button className="admin-choice-card" onClick={() => setPage('activitiesEquipment')}>
                <ClipboardList size={34} />
                <div><h3>Activity Library</h3><p>Manage SOP-led staffing, qualifications, equipment, restrictions and safety rules for every activity.</p></div>
              </button>}
              {canManageStaff && <button className="admin-choice-card" onClick={() => setPage('staff')}>
                <Users size={34} />
                <div><h3>Staff</h3><p>Manage staff accounts, roles and availability.</p></div>
              </button>}
              {!isAdminAccount && <button className="admin-choice-card" onClick={() => setPage('signoffs')}>
                <ShieldCheck size={34} />
                <div><h3>Sign-off</h3><p>Search staff and manage activity sign-offs.</p></div>
              </button>}
              {canViewLogs && <button className="admin-choice-card" onClick={() => setPage('logs')}>
                <ClipboardList size={34} />
                <div><h3>Logs</h3><p>Review water-lead permission confirmations.</p></div>
              </button>}
              {canViewLogs && <button className="admin-choice-card" onClick={() => setPage('staffingLogs')}><FileSpreadsheet size={34}/><div><h3>Staffing Logs</h3><p>Locked weekly staffing records and historical downloads.</p></div></button>}
              {canManageStaff && <section className="display-manager-card admin-account-card"><div><Users size={34}/><div><h3>Admin Account</h3><p>Give an existing signed-up account access to Admin → Programme Files only.</p></div></div><div className="admin-account-form"><input type="email" value={adminAccountEmail} onChange={(event) => setAdminAccountEmail(event.target.value)} placeholder="admin@school.co.uk"/><button className="primary" disabled={adminAccountBusy || !adminAccountEmail.trim()} onClick={() => void makeAdminAccount()}>{adminAccountBusy ? 'Saving…' : 'Give Admin Access'}</button></div></section>}
              {canManageStaff && <button className="admin-choice-card" onClick={() => setPage('formerStaff')}><History size={34}/><div><h3>Former Staff</h3><p>Employment start and leaving records.</p></div></button>}
              {canManageStaff && <button className="admin-choice-card" onClick={() => setPage('loanHistory')}><Users size={34}/><div><h3>Loan Staff History</h3><p>Reactivate loan staff or add them permanently.</p></div></button>}
            </section>
          </Panel>
        )}

        {page === 'activitiesEquipment' && canManageStaff && (
          <Panel title="Activity Library" onBack={() => setPage('admin')}>
            <section className="activity-equipment-intro">
              <div><p className="eyebrow">SOP-driven operations engine</p><h3>Staffing, qualifications, equipment and safety controls</h3><p>These rules drive Programme Builder, Auto-fill and compliance warnings. Shared activities from the same school and session are combined automatically.</p></div>
              <div className="capacity-example"><strong>Compliance status</strong><span>Rules stored per activity</span><span>SOP and RA references visible to managers</span></div>
            </section>
            <section className="add-activity-strip">
              <input value={newActivityCode} onChange={(event) => setNewActivityCode(event.target.value.toUpperCase())} placeholder="Code, e.g. RAFT" />
              <input value={newActivityName} onChange={(event) => setNewActivityName(event.target.value)} placeholder="Activity name" />
              <button className="primary" onClick={addActivity}><Plus size={17}/>Add activity</button>
            </section>
            <div className="activity-equipment-list">
              {activities.filter((activity) => activity.code !== 'Z').map((activity) => (
                <article className={`activity-equipment-card ${activity.enabled === false ? 'disabled' : ''}`} key={activity.code}>
                  <div className="activity-equipment-title">
                    <span className="activity-colour-dot" style={{ background: activity.colour ?? '#dce8f5' }} />
                    <div><strong>{activity.code}</strong><input value={activity.name} onChange={(event) => updateActivityDefinition(activity.code, { name: event.target.value })}/></div>
                    <label className="activity-enabled"><input type="checkbox" checked={activity.enabled !== false} onChange={(event) => updateActivityDefinition(activity.code, { enabled: event.target.checked })}/>Enabled</label>
                  </div>
                  <div className="activity-equipment-fields">
                    <label>Activity colour<input type="color" value={activity.colour ?? '#dce8f5'} onChange={(event) => updateActivityDefinition(activity.code, { colour: event.target.value })}/></label>
                    <label>Equipment quantity<input type="number" min="0" value={activity.equipmentQuantity ?? 0} onChange={(event) => updateActivityDefinition(activity.code, { equipmentQuantity: Math.max(0, Number(event.target.value) || 0) })}/></label>
                    <label>Maximum groups per session<input type="number" min="1" value={activity.capacity ?? 1} onChange={(event) => updateActivityDefinition(activity.code, { capacity: Math.max(1, Number(event.target.value) || 1) })}/></label>
                    <label>Staffing method<select value={activity.staffingRuleType ?? 'per_group'} onChange={(event) => updateActivityDefinition(activity.code, { staffingRuleType: event.target.value as Activity['staffingRuleType'] })}><option value="per_group">1 instructor per group</option><option value="per_x_groups">1 instructor per X groups</option><option value="fixed">Fixed number of instructors</option><option value="manual">Manual / one per group</option></select></label>
                    <label>{activity.staffingRuleType === 'fixed' ? 'Fixed instructors' : activity.staffingRuleType === 'per_x_groups' ? 'Groups per instructor' : 'Staffing value'}<input type="number" min="1" disabled={activity.staffingRuleType === 'per_group' || activity.staffingRuleType === 'manual'} value={activity.staffingRuleValue ?? 1} onChange={(event) => updateActivityDefinition(activity.code, { staffingRuleValue: Math.max(1, Number(event.target.value) || 1) })}/></label>
                    <label>Staffing priority<select value={activity.staffingPriority ?? 2} onChange={(event) => updateActivityDefinition(activity.code, { staffingPriority: Number(event.target.value) })}><option value="1">1 · Highest</option><option value="2">2 · Normal</option><option value="3">3 · Shared / lowest</option></select></label>
                    <label>Minimum instructors<input type="number" min="1" value={activity.minimumInstructors ?? 1} onChange={(event) => updateActivityDefinition(activity.code, { minimumInstructors: Math.max(1, Number(event.target.value) || 1) })}/></label>
                    <label>Maximum participants per group<input type="number" min="0" value={activity.maximumGroupSize ?? 0} onChange={(event) => updateActivityDefinition(activity.code, { maximumGroupSize: Math.max(0, Number(event.target.value) || 0) })} placeholder="0 = not set"/></label>
                    <label>Required qualification<input value={activity.requiredQualifications ?? ''} onChange={(event) => updateActivityDefinition(activity.code, { requiredQualifications: event.target.value })} placeholder="Centre sign-off / NGB qualification"/></label>
                    <label>Lead qualification<input value={activity.leadQualification ?? ''} onChange={(event) => updateActivityDefinition(activity.code, { leadQualification: event.target.value })} placeholder="Optional lead requirement"/></label>
                    <label className="activity-enabled"><input type="checkbox" checked={activity.shareable === true} onChange={(event) => updateActivityDefinition(activity.code, { shareable: event.target.checked })}/>Can be shared across groups</label>
                    <label className="activity-enabled"><input type="checkbox" checked={activity.safetyBoatRequired === true} onChange={(event) => updateActivityDefinition(activity.code, { safetyBoatRequired: event.target.checked })}/>Safety boat required</label>
                    <label className="activity-notes">Weather / environmental restrictions<textarea value={activity.weatherRestrictions ?? ''} onChange={(event) => updateActivityDefinition(activity.code, { weatherRestrictions: event.target.value })} placeholder="Wind, visibility, flooding, lightning or seasonal restrictions"/></label>
                    <label className="activity-notes">Setup checklist<textarea value={activity.setupChecklist ?? ''} onChange={(event) => updateActivityDefinition(activity.code, { setupChecklist: event.target.value })} placeholder="One item per line"/></label>
                    <label className="activity-notes">Pack-away checklist<textarea value={activity.packAwayChecklist ?? ''} onChange={(event) => updateActivityDefinition(activity.code, { packAwayChecklist: event.target.value })} placeholder="One item per line"/></label>
                    <label className="activity-notes">Emergency procedure<textarea value={activity.emergencyProcedure ?? ''} onChange={(event) => updateActivityDefinition(activity.code, { emergencyProcedure: event.target.value })}/></label>
                    <label>SOP reference<input value={activity.sopReference ?? ''} onChange={(event) => updateActivityDefinition(activity.code, { sopReference: event.target.value })}/></label>
                    <label>Risk assessment reference<input value={activity.riskAssessmentReference ?? ''} onChange={(event) => updateActivityDefinition(activity.code, { riskAssessmentReference: event.target.value })}/></label>
                    <label className="activity-notes">Manager notes<input value={activity.notes ?? ''} onChange={(event) => updateActivityDefinition(activity.code, { notes: event.target.value })} placeholder="Optional operational note"/></label>
                  </div>
                  <div className="activity-equipment-footer"><span>{activity.equipmentQuantity ?? 0} equipment recorded · {activity.capacity ?? 1} group{(activity.capacity ?? 1) === 1 ? '' : 's'} at once · {activity.staffingRuleType === 'per_x_groups' ? `1 instructor per ${activity.staffingRuleValue ?? 1} groups` : activity.staffingRuleType === 'fixed' ? `${activity.staffingRuleValue ?? 1} instructors fixed` : '1 instructor per group'} · {activity.requiredQualifications ? 'Qualification rule set' : 'Qualification rule needed'} · {activity.sopReference ? 'SOP linked' : 'SOP reference needed'}</span><button className="remove-button" onClick={() => deleteActivity(activity.code)}><Trash2 size={16}/>Remove</button></div>
                </article>
              ))}
            </div>
          </Panel>
        )}

        {page === 'staff' && canManageStaff && (
          <Panel title="Staff management" onBack={() => setPage('admin')}>
            <div className="staff-page-toolbar">
              <div>
                <p>
                  Add new instructors, set team leaders and review how many
                  activities each person is signed off to run.
                </p>
              </div>
              <div className="staff-toolbar-actions"><button className="primary" onClick={() => { setAddingLoanStaff(false); setShowAddStaff(true) }}><Plus size={18}/>Add staff</button><button className="secondary-action" onClick={() => { setAddingLoanStaff(true); setShowAddStaff(true) }}><Plus size={18}/>Add loan staff</button></div>
            </div>

            {showAddStaff && (
              <section className="add-staff-panel">
                <div className="add-staff-heading">
                  <div>
                    <p className="eyebrow">Manager controls</p>
                    <h3>{addingLoanStaff ? 'Add loan staff' : 'Add staff member'}</h3>
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

                <label>Start date</label><input type="date" value={newStaffStartDate} onChange={(event)=>setNewStaffStartDate(event.target.value)}/>{addingLoanStaff && <><label>Expected loan end date</label><input type="date" value={newLoanEndDate} onChange={(event)=>setNewLoanEndDate(event.target.value)}/></>}

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
                  {addingLoanStaff ? 'Add loan staff' : 'Add staff member'}
                </button>
              </section>
            )}

            <div className="staff-management-list">
              {staff.map((member) => (
                <article className="staff-management-card" key={member.id}>
                  <div>
                    <div className="staff-name-line">
                      <h3>{member.name}</h3>{member.employmentType === 'loan' && <span className="loan-badge">Loan staff</span>}<small className="staff-code">{member.staffCode ?? 'Legacy record'}</small>
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
                    {member.employmentType === 'loan' ? <button className="delete-staff" onClick={() => endLoan(member.id)}><History size={16}/>End Loan</button> : <button className="delete-staff" onClick={() => markLeftCompany(member.id)}><History size={16}/>Left Company</button>}
                  </div>
                </article>
              ))}
            </div>
          </Panel>
        )}


        {page === 'staffingLogs' && canViewLogs && (
          <Panel title={selectedStaffingLogMonth ? staffingMonthLabel(selectedStaffingLogMonth) : 'Staffing Logs'} onBack={() => selectedStaffingLogMonth ? setSelectedStaffingLogMonth('') : setPage('admin')}>
            <section className="staffing-log-intro"><div><p className="eyebrow">Permanent weekly records</p><h3>{selectedStaffingLogMonth ? `Weeks in ${staffingMonthLabel(selectedStaffingLogMonth)}` : 'Archived staffing months'}</h3><p>Archived weeks are locked snapshots and do not change when a later programme is edited or uploaded.</p></div>{programme && !selectedStaffingLogMonth && <button className="primary" onClick={() => archiveSnapshot(programme, true)}><History size={17}/>Archive current week</button>}</section>
            {!selectedStaffingLogMonth ? (
              <div className="staffing-month-grid">
                {staffingArchives.length === 0 ? <p>No staffing weeks have been archived yet.</p> : Array.from(new Set(staffingArchives.map(staffingArchiveMonth))).sort().reverse().map((monthKey) => {
                  const count = staffingArchives.filter((archive) => staffingArchiveMonth(archive) === monthKey).length
                  return <button className="staffing-month-card" key={monthKey} onClick={() => setSelectedStaffingLogMonth(monthKey)}><History size={25}/><span><strong>{staffingMonthLabel(monthKey)}</strong><small>{count} archived week{count === 1 ? '' : 's'}</small></span><span className="staffing-month-open">Open</span></button>
                })}
              </div>
            ) : (
              <div className="history-list">{staffingArchives.filter((archive) => staffingArchiveMonth(archive) === selectedStaffingLogMonth).sort((a,b) => b.archivedAt.localeCompare(a.archivedAt)).map((archive) => {
                const days = Array.from(new Set(archive.programme.rows.map((row) => row.day)))
                return <article className="history-card staffing-log-card" key={archive.id}><div><h3>{archive.weekKey}</h3><p>{archive.title} · {archive.sourceFileName}</p><small>Archived {new Date(archive.archivedAt).toLocaleString('en-GB')} by {archive.archivedBy}</small></div><div className="staffing-log-actions"><button className="primary" onClick={() => exportArchivedStaffing(archive)}><FileSpreadsheet size={16}/>Download full week</button>{days.map((day) => <button className="secondary-action" key={day} onClick={() => exportArchivedStaffing(archive, day)}>{day}</button>)}</div></article>
              })}</div>
            )}
          </Panel>
        )}


        {page === 'formerStaff' && canManageStaff && (<Panel title="Former Staff" onBack={() => setPage('admin')}><div className="history-list">{formerStaff.length === 0 ? <p>No former staff records yet.</p> : formerStaff.map((record)=><article className="history-card" key={record.member.id}><div><h3>{record.member.name}</h3><p>{roleLabel(resolvedRole(record.member))} · {record.member.staffCode ?? 'Legacy record'}</p><p>Started: {record.member.startDate ?? 'Not recorded'} · Left: {record.endDate}</p>{record.notes && <p>{record.notes}</p>}</div><button className="primary" onClick={()=>reinstateFormer(record)}>Reinstate Staff</button></article>)}</div></Panel>)}

        {page === 'loanHistory' && canManageStaff && (<Panel title="Loan Staff History" onBack={() => setPage('admin')}><div className="history-list">{loanHistory.length === 0 ? <p>No completed loan staff records yet.</p> : loanHistory.map((record)=><article className="history-card" key={record.member.id}><div><h3>{record.member.name}</h3><p>{roleLabel(resolvedRole(record.member))} · {record.member.staffCode ?? 'Legacy record'}</p><p>{record.loanPeriods?.length ?? 1} loan period{(record.loanPeriods?.length ?? 1)===1?'':'s'}</p>{record.loanPeriods?.map((period,index)=><small key={index}>Loan {index+1}: {period.startDate} to {period.endDate}{period.notes?` — ${period.notes}`:''}</small>)}</div><div className="history-actions"><button className="secondary-action" onClick={()=>reactivateLoan(record)}>Reactivate as Loan Staff</button><button className="primary" onClick={()=>convertLoanToPermanent(record)}>Add to My Centre</button></div></article>)}</div></Panel>)}

        {page === 'holidays' && (
          <Panel title="Days Off & Sickness" onBack={() => setPage('staffing')}>
            <div className="days-off-toolbar no-print">
              <div className="staffing-view-toggle">
                <button className={daysOffView === 'month' ? 'active' : ''} onClick={() => setDaysOffView('month')}>Month View</button>
                <button className={daysOffView === 'week' ? 'active' : ''} onClick={() => setDaysOffView('week')}>Weekly View</button>
                <button className={daysOffView === 'day' ? 'active' : ''} onClick={() => setDaysOffView('day')}>Daily View</button>
              </div>
              <div className="days-off-export-actions">
                <button className="secondary-action" onClick={() => setShowDaysOffHelp(true)}>? Help & Key</button>
                <button className="secondary-action" onClick={() => downloadDaysOffExcel(daysOffView)}><FileSpreadsheet size={17}/>Download Excel</button>
              </div>
            </div>

            {canManageHolidays && <section className="days-off-range-editor no-print">
              <select value={daysOffStaffId} onChange={(e) => setDaysOffStaffId(e.target.value)}><option value="">Select staff member</option>{staff.map(m=><option key={m.id} value={m.id}>{m.name}</option>)}</select>
              <select value={daysOffStatus} onChange={(e)=>setDaysOffStatus(e.target.value as DayOffStatus)}>
                <option value="off">OFF</option><option value="hol">HOL</option><option value="sick">SICK</option><option value="am_off">AM OFF</option><option value="pm_off">PM OFF</option>
              </select>
              <input type="date" value={daysOffStart} onChange={(e)=>setDaysOffStart(e.target.value)}/>
              <input type="date" value={daysOffEnd} onChange={(e)=>setDaysOffEnd(e.target.value)}/>
              <button className="primary" onClick={saveDaysOffRange}>Set dates</button>
            </section>}

            {daysOffView === 'month' ? <section className="print-month-sheet">
              <div className="holiday-summary">
                <div><p className="eyebrow">Monthly Days Off</p><h3>{holidayMonth.toLocaleDateString('en-GB',{month:'long',year:'numeric'})}</h3></div>
                <div className="holiday-month-actions no-print"><button onClick={()=>setHolidayMonth(new Date(holidayMonth.getFullYear(),holidayMonth.getMonth()-1,1))}>Previous</button><button onClick={()=>setHolidayMonth(new Date())}>Today</button><button onClick={()=>setHolidayMonth(new Date(holidayMonth.getFullYear(),holidayMonth.getMonth()+1,1))}>Next</button></div>
              </div>
              <div className="holiday-weekdays">{['Mon','Tue','Wed','Thu','Fri','Sat','Sun'].map(d=><strong key={d}>{d}</strong>)}</div>
              <div className="holiday-calendar">{holidayCalendarDays().map(date=>{const key=dateKey(date); const entries=daysOff.filter(x=>x.day===key); const legacyHol=holidays.filter(h=>h.start_date<=key&&h.end_date>=key).map(h=>({id:h.id,staff_name:h.staff_name,status:'hol' as DayOffStatus})); return <article key={key} className={`holiday-day ${date.getMonth()!==holidayMonth.getMonth()?'outside':''}`}><span className="holiday-date">{date.getDate()}</span>{[...legacyHol,...entries].sort((a:any,b:any)=>{const ar=staff.find(m=>m.id===a.staff_id||normaliseIdentity(m.name)===normaliseIdentity(a.staff_name));const br=staff.find(m=>m.id===b.staff_id||normaliseIdentity(m.name)===normaliseIdentity(b.staff_name));const rank:Record<StaffRole,number>={centreManager:0,activityManager:1,teamLeader:2,staff:3};return (ar?rank[resolvedRole(ar)]:4)-(br?rank[resolvedRole(br)]:4)||String(a.staff_name).localeCompare(String(b.staff_name))}).map((x:any)=><div key={`${x.id}-${key}`} className={`day-off-entry status-${x.status}`}><span>{x.staff_name}</span><b>{x.note==='blank-red'?'':dayOffLabel(x.status)}</b></div>)}</article>})}</div>
            </section> : daysOffView === 'week' ? <section className="weekly-days-off print-week-sheet">
              <div className="weekly-sheet-heading"><div><h2>WEEKLY STAFFING</h2><p>Days Off</p></div><div className="week-nav no-print"><button onClick={()=>{const d=new Date(daysOffWeek);d.setDate(d.getDate()-7);setDaysOffWeek(d)}}>Previous</button><button onClick={()=>{const n=new Date();n.setDate(n.getDate()-((n.getDay()+6)%7));setDaysOffWeek(n)}}>This week</button><button onClick={()=>{const d=new Date(daysOffWeek);d.setDate(d.getDate()+7);setDaysOffWeek(d)}}>Next</button></div></div>
              <div className="weekly-grid" style={{gridTemplateColumns:`180px repeat(7,minmax(105px,1fr))`}}>
                <div className="weekly-head staff-head">Staff</div>{daysOffWeekDates().map(d=><div className="weekly-head" key={dateKey(d)}><strong>{d.toLocaleDateString('en-GB',{weekday:'long'})}</strong><span>{d.toLocaleDateString('en-GB',{day:'numeric',month:'short'})}</span></div>)}
                {sortedDaysOffStaff().map((member,rowIndex)=><Fragment key={member.id}><div className="staff-name-cell"><span>{member.name}</span><small>{roleLabel(resolvedRole(member))}</small></div>{daysOffWeekDates().map((d,columnIndex)=>{const key=dateKey(d); const cellKey=`${member.id}-${key}`; const existing=daysOff.find(x=>x.staff_id===member.id&&x.day===key); return <div ref={node=>{if(node) weeklyCellRefs.current.set(cellKey,node); else weeklyCellRefs.current.delete(cellKey)}} tabIndex={0} role="gridcell" aria-label={`${member.name}, ${d.toLocaleDateString('en-GB')}`} className={`day-off-cell keyboard-cell ${selectedDaysOffCell===cellKey?'selected-cell':''} ${existing?`status-${existing.status}`:''} ${existing?.note==='blank-red'?'blank-red-cell':''}`} key={cellKey} onFocus={()=>setSelectedDaysOffCell(cellKey)} onKeyDown={event=>handleWeeklyCellKeyDown(event,member,key,rowIndex,columnIndex)}><select tabIndex={-1} className="no-print" value={existing?.status??'working'} disabled={!canManageHolidays && accountRole!=='teamLeader'} onChange={e=>setSingleDayOff(member,key,e.target.value as DayOffStatus|'working')}><option value="working">Working</option>{canManageHolidays&&<option value="off">OFF</option>}{canManageHolidays&&<option value="hol">HOL</option>}<option value="sick">SICK</option>{canManageHolidays&&<option value="am_off">AM OFF</option>}{canManageHolidays&&<option value="pm_off">PM OFF</option>}</select><strong className="print-only">{daysOffDisplayLabel(existing)}</strong></div>})}</Fragment>)}
              </div>
              <aside className="keyboard-key no-print" aria-label="Weekly staffing keyboard key">
                <div className="keyboard-key-heading"><strong>Keyboard key</strong><span>Click a cell first</span></div>
                <div className="keyboard-key-grid">
                  <span><kbd>← ↑ → ↓</kbd> Move</span>
                  <span className="key-off"><kbd>O</kbd> OFF</span>
                  <span className="key-hol"><kbd>H</kbd> Holiday</span>
                  <span className="key-sick"><kbd>S</kbd> Sick with text</span>
                  <span className="key-sick"><kbd>R</kbd> Blank red sick box</span>
                  <span className="key-am"><kbd>A</kbd> AM OFF</span>
                  <span className="key-pm"><kbd>P</kbd> PM OFF</span>
                  <span><kbd>Delete</kbd> Clear / Working</span>
                  <span><kbd>Enter</kbd> Move down</span>
                  <span><kbd>Tab</kbd> Move right</span>
                </div>
              </aside>
              <div className="days-off-legend"><span className="status-off">OFF</span><span className="status-hol">HOL</span><span className="status-sick">SICK</span><span className="status-am_off">AM OFF</span><span className="status-pm_off">PM OFF</span></div>
            </section> : <section className="daily-days-off print-day-sheet">
              <div className="weekly-sheet-heading"><div><p className="eyebrow">Daily Days Off</p><h2>{parseDateKey(daysOffDay).toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long',year:'numeric'})}</h2></div><div className="week-nav no-print"><button onClick={()=>{const d=parseDateKey(daysOffDay);d.setDate(d.getDate()-1);setDaysOffDay(dateKey(d))}}>Previous</button><button onClick={()=>setDaysOffDay(dateKey(new Date()))}>Today</button><button onClick={()=>{const d=parseDateKey(daysOffDay);d.setDate(d.getDate()+1);setDaysOffDay(dateKey(d))}}>Next</button><input type="date" value={daysOffDay} onChange={e=>setDaysOffDay(e.target.value)}/></div></div>
              <div className="daily-days-off-grid">
                <div className="daily-head">Staff</div><div className="daily-head">Status</div><div className="daily-head">Availability</div>
                {sortedDaysOffStaff().map((member,rowIndex)=>{const existing=daysOff.find(x=>x.staff_id===member.id&&x.day===daysOffDay);const availability=!existing?'Available all day':existing.status==='am_off'?'Unavailable Sessions 1 & 2':existing.status==='pm_off'?'Unavailable Session 5':'Unavailable all day';const cellKey=`${member.id}-${daysOffDay}`;return <Fragment key={member.id}><div className="staff-name-cell"><span>{member.name}</span><small>{roleLabel(resolvedRole(member))}</small></div><div ref={node=>{if(node) dailyCellRefs.current.set(cellKey,node); else dailyCellRefs.current.delete(cellKey)}} tabIndex={0} role="gridcell" aria-label={`${member.name}, ${parseDateKey(daysOffDay).toLocaleDateString('en-GB')}`} className={`day-off-cell keyboard-cell ${selectedDaysOffCell===cellKey?'selected-cell':''} ${existing?`status-${existing.status}`:''} ${existing?.note==='blank-red'?'blank-red-cell':''}`} onFocus={()=>setSelectedDaysOffCell(cellKey)} onKeyDown={event=>handleDailyCellKeyDown(event,member,rowIndex)}><select tabIndex={-1} className="no-print" value={existing?.status??'working'} disabled={!canManageHolidays&&accountRole!=='teamLeader'} onChange={e=>setSingleDayOff(member,daysOffDay,e.target.value as DayOffStatus|'working')}><option value="working">Working</option>{canManageHolidays&&<option value="off">OFF</option>}{canManageHolidays&&<option value="hol">HOL</option>}<option value="sick">SICK</option>{canManageHolidays&&<option value="am_off">AM OFF</option>}{canManageHolidays&&<option value="pm_off">PM OFF</option>}</select><strong className="print-only">{daysOffDisplayLabel(existing)}</strong></div><div className="daily-availability">{availability}</div></Fragment>})}
              </div>
              <div className="days-off-legend"><span className="status-off">OFF</span><span className="status-hol">HOL</span><span className="status-sick">SICK</span><span className="status-am_off">AM OFF</span><span className="status-pm_off">PM OFF</span></div>
            </section>}

            {canManageHolidays && <section className="payroll-download-panel no-print">
              <div><h3>Monthly payroll</h3><p>Uses the Jun payroll layout and adds each new month to the same annual Excel workbook. New staff are calculated from their employment start date to month end.</p></div>
              <div className="payroll-actions"><button className="secondary-action" onClick={payrollSync}><History size={17}/>Payroll Sync</button><button className="primary" onClick={downloadPayroll}><FileSpreadsheet size={17}/>Download Payroll</button>{payrollSyncAt && <small>Last synced {new Date(payrollSyncAt).toLocaleString('en-GB')}</small>}</div>
            </section>}

            {showDaysOffHelp && <div className="days-off-help-backdrop no-print" role="presentation" onMouseDown={(event)=>{if(event.target===event.currentTarget)setShowDaysOffHelp(false)}}>
              <section className="days-off-help-dialog" role="dialog" aria-modal="true" aria-labelledby="days-off-help-title">
                <div className="days-off-help-header"><div><p className="eyebrow">Days Off</p><h2 id="days-off-help-title">Help and keyboard key</h2></div><button type="button" className="secondary-action" onClick={()=>setShowDaysOffHelp(false)}>Close</button></div>
                <h3>Weekly keyboard shortcuts</h3>
                <div className="keyboard-key-grid help-key-grid">
                  <span><kbd>← ↑ → ↓</kbd> Move between cells</span><span className="key-off"><kbd>O</kbd> OFF — unavailable all day</span><span className="key-hol"><kbd>H</kbd> HOL — unavailable all day</span><span className="key-sick"><kbd>S</kbd> SICK — shows SICK</span><span className="key-sick"><kbd>R</kbd> SICK — blank red box</span><span className="key-am"><kbd>A</kbd> AM OFF — blocks Sessions 1 and 2</span><span className="key-pm"><kbd>P</kbd> PM OFF — blocks Session 5</span><span><kbd>Delete</kbd> or <kbd>Backspace</kbd> Clear to Working</span><span><kbd>Enter</kbd> Move down</span><span><kbd>Tab</kbd> Move right</span>
                </div>
                <h3>Using the grid</h3><p>Click a weekly cell, then press a shortcut key. The outlined cell is the selected cell. The arrow keys move around the week without using the mouse.</p>
                <h3>Excel</h3><p>Download Excel creates an editable workbook using the selected Month, Weekly or Daily view. Open it in Excel to edit or print it.</p>
              </section>
            </div>}

            <section className="staff-work-summary no-print">
              <div className="staff-work-summary-heading"><div><p className="eyebrow">Staff overview</p><h3>Work and absence summary</h3></div><select value={holidaySummaryStaffId} onChange={(e)=>setHolidaySummaryStaffId(e.target.value)}><option value="">Select staff member</option>{staff.map(m=><option key={m.id} value={m.id}>{m.name}</option>)}</select></div>
              {!holidayStaffSummary?<p className="empty-summary">Select a staff member to view their totals.</p>:<div className="staff-stat-grid"><article><strong>{holidayStaffSummary.daysWorked}</strong><span>Days worked</span></article><article><strong>{holidayStaffSummary.sessionsWorked}</strong><span>Total sessions</span></article><article><strong>{holidayStaffSummary.holidayDays}</strong><span>Holiday days</span></article><article><strong>{holidayStaffSummary.sickDays}</strong><span>Sick days</span></article><article className="wide"><strong>{holidayStaffSummary.mostRun?`${holidayStaffSummary.mostRun[0]} — ${holidayStaffSummary.mostRun[1]}`:'No activity data yet'}</strong><span>Most-run activity</span></article><article className="wide"><strong>{holidayStaffSummary.leastRun?`${holidayStaffSummary.leastRun[0]} — ${holidayStaffSummary.leastRun[1]}`:'No activity data yet'}</strong><span>Least-run activity</span></article></div>}
            </section>
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

      {lastSharedUpdate && (
        <footer className="last-updated-footer" title={new Date(lastSharedUpdate.updated_at).toLocaleString('en-GB')}>
          Last updated by {lastSharedUpdate.updated_by_name || lastSharedUpdate.updated_by_email} · {new Date(lastSharedUpdate.updated_at).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })}
        </footer>
      )}

      {pendingWaterConfirmation && (() => {
        const member = staff.find((item) => item.id === pendingWaterConfirmation.staffId)
        return (
          <div className="modal-backdrop">
            <section className="water-confirmation-modal" role="dialog" aria-modal="true">
              <CircleAlert size={38} />
              <h2>Water staffing confirmation</h2>
              <p>There is no spare qualified {pendingWaterConfirmation.discipline === 'canoe' ? 'Canoe Lead' : 'Kayak Lead'} for {pendingWaterConfirmation.day}, Session {pendingWaterConfirmation.session}.</p>
              <p><strong>{member?.name}</strong> will run G{pendingWaterConfirmation.leadGroup} and oversee {pendingWaterConfirmation.overseenGroups.map((group) => `G${group}`).join(', ')}.</p>
              <p>Have you spoken to the Head of Centre or Activities Manager and received permission?</p>
              <div className="confirmation-actions">
                <button onClick={() => setPendingWaterConfirmation(null)}>No, go back</button>
                <button className="primary" onClick={() => confirmWaterLeadException('Activities Manager')}>Yes — Activities Manager</button>
                <button className="primary" onClick={() => confirmWaterLeadException('Head of Centre')}>Yes — Head of Centre</button>
              </div>
            </section>
          </div>
        )
      })()}

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
  const rows = useMemo(() => {
    const merged = new Map<string, ProgrammeRow>()
    for (const row of programme.rows) {
      const key = `${row.day}|${row.session}`
      const current = merged.get(key)
      if (!current) {
        merged.set(key, { ...row, id: `centre-${row.day}-${row.session}`, cells: row.cells.map((cell) => ({ ...cell })) })
        continue
      }
      const cells = new Map(current.cells.map((cell) => [cell.group, cell]))
      row.cells.forEach((cell) => {
        const existing = cells.get(cell.group)
        if (!existing?.activityCode || cell.activityCode) cells.set(cell.group, { ...cell })
      })
      const labels = [current.schoolLabel, row.schoolLabel].filter(Boolean).flatMap((value) => String(value).split(/\s*\/\s*/)).filter((value, index, list) => list.indexOf(value) === index)
      merged.set(key, { ...current, schoolLabel: labels.join(' / '), cells: Array.from(cells.values()).sort((a, b) => a.group - b.group) })
    }
    return Array.from(merged.values()).sort((a, b) => {
      const dayDiff = weekdayRank(a.day) - weekdayRank(b.day)
      return dayDiff || Number(a.session) - Number(b.session)
    })
  }, [programme])

  return (
    <div className="programme-scroll centre-grid-scroll">
      <table className="programme-table centre-programme-table">
        <thead>
          <tr>
            <th className="sticky-day">Day</th>
            <th className="sticky-session">Ses</th>
            {programme.groupNumbers.map((group) => (
              <th key={group}><span>G{group}</span><small>{programme.schoolDetails?.find((school) => school.groupNumbers?.includes(group))?.schoolName ?? ''}</small></th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, index) => {
            const previous = rows[index - 1]
            const showDay = !previous || previous.day !== row.day
            return (
              <tr key={row.id} className={showDay ? 'programme-day-start' : ''}>
                <th className="sticky-day">{showDay ? row.day : ''}</th>
                <th className="sticky-session">{row.session}</th>
                {programme.groupNumbers.map((group) => {
                  const cell = row.cells.find((item) => item.group === group)
                  const code = cell?.activityCode ?? ''
                  const schoolName = programme.schoolDetails?.find((school) => school.groupNumbers?.includes(group))?.schoolName ?? ''
                  const isArrival = row.session === '3' && schoolName && code.toLowerCase() === schoolName.toLowerCase()
                  const display = isArrival ? schoolName.toUpperCase() : code || '—'
                  return (
                    <td key={group}>
                      <button
                        className={`programme-cell code-${code.toLowerCase().replace(/[^a-z0-9]+/g, '-')} ${isArrival ? 'arrival-cell' : ''}`}
                        onClick={() => {
                          const sourceRow = programme.rows.find((candidate) => candidate.day === row.day && candidate.session === row.session && candidate.cells.some((item) => item.group === group))
                          onSelect(sourceRow ?? row, group)
                        }}
                        title={isArrival ? `${schoolName} arrival` : code ? activityNameFromList(activities, code) : 'Empty'}
                      >
                        {display}
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
  onBack?: () => void
  children: React.ReactNode
}) {
  return (
    <section className="panel">
      {onBack && <button className="back" onClick={onBack}>
        <ChevronLeft size={18} />
        Back
      </button>}
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
