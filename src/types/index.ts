export type Page = 'dashboard' | 'programme' | 'arrivals' | 'staffing' | 'schoolNotes' | 'admin' | 'staff' | 'holidays' | 'signoffs' | 'logs' | 'formerStaff' | 'loanHistory' | 'staffingLogs' | 'programmeBuilder' | 'activitiesEquipment' | 'programmeArchive'

export type Activity = {
  code: string
  name: string
  colour?: string
  equipmentQuantity?: number
  capacity?: number
  enabled?: boolean
  notes?: string
  staffingRuleType?: 'per_group' | 'per_x_groups' | 'fixed' | 'manual'
  staffingRuleValue?: number
  staffingPriority?: number
  requiredQualifications?: string
  leadQualification?: string
  minimumInstructors?: number
  maximumGroupSize?: number
  safetyBoatRequired?: boolean
  shareable?: boolean
  weatherRestrictions?: string
  setupChecklist?: string
  packAwayChecklist?: string
  emergencyProcedure?: string
  sopReference?: string
  riskAssessmentReference?: string
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
  startDate?: string
  endDate?: string
  schoolDetails?: { id: string; schoolName: string; programmeName: string; purchaseType: 'normal' | 'bargain' | 'super' | 'outdoor'; arrivalDate: string; departureDate: string; notes: string; groupNumbers?: number[] }[]
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
