export type Page = 'dashboard' | 'programme' | 'arrivals' | 'staffing' | 'schoolNotes' | 'admin' | 'staff' | 'holidays' | 'signoffs' | 'logs' | 'formerStaff' | 'loanHistory'

export type Activity = {
  code: string
  name: string
}

export type ProgrammeCell = {
  group: number
  activityCode: string
}

export type ProgrammeRow = {
  id: string
  day: string
  session: string
  schoolLabel?: string
  cells: ProgrammeCell[]
}

export type ProgrammeImport = {
  title: string
  sheetName: string
  groupNumbers: number[]
  rows: ProgrammeRow[]
  importedAt: string
  sourceFileName: string
}

export type StaffRole =
  | 'staff'
  | 'teamLeader'
  | 'activityManager'
  | 'centreManager'

export type StaffMember = {
  id: string
  name: string
  email?: string
  qualifications: string[]
  signOffs?: Record<string, string>
  role?: StaffRole
  teamLeader?: boolean
  staffCode?: string
  employmentType?: 'permanent' | 'loan'
  startDate?: string
  loanEndDate?: string
}

export type ArchivedStaff = {
  member: StaffMember
  archivedAt: string
  endDate: string
  notes?: string
  archiveType: 'former' | 'loan'
  loanPeriods?: { startDate: string; endDate: string; notes?: string }[]
}

export type StaffingAssignment = Record<string, string>


export type ArrivalAssignment = {
  leaderId?: string
  guideIds: string[]
  flatIds?: string[]
  notes?: string
  noteStaffIds?: string[]
}
