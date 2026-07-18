import { useEffect, useMemo, useState } from 'react'
import { Building2, CalendarDays, CloudSun, HeartPulse, School, Users } from 'lucide-react'
import { supabase } from './lib/supabase'
import type { ArrivalAssignment, ProgrammeImport, ProgrammeRow, StaffMember } from './types'

type DayOffEntry = {
  staff_id: string
  staff_name: string
  day: string
  status: 'off' | 'hol' | 'sick' | 'am_off' | 'pm_off'
}

type LiveState = {
  programme: ProgrammeImport | null
  staff: StaffMember[]
  arrivalAssignments: Record<string, ArrivalAssignment>
}

type WeatherState = {
  temperature: number
  wind: number
  code: number
}

const buildings = ['Kingfisher', 'Swan', 'Grebe', 'Bittern', 'Mallard', 'Teal']

function isoDate(date: Date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`
}

function weekdayRank(value: string) {
  const key = value.toUpperCase().replace(/[^A-Z]/g, '')
  return ({ MON: 1, MONDAY: 1, TUE: 2, TUES: 2, TUESDAY: 2, WED: 3, WEDNESDAY: 3, THU: 4, THUR: 4, THURS: 4, THURSDAY: 4, FRI: 5, FRIDAY: 5, SAT: 6, SATURDAY: 6, SUN: 7, SUNDAY: 7 } as Record<string, number>)[key] ?? 99
}

function dateForDay(programme: ProgrammeImport, day: string) {
  if (/^\d{4}-\d{2}-\d{2}$/.test(day)) return day
  if (!programme.startDate) return ''
  const start = new Date(`${programme.startDate}T12:00:00`)
  const rank = weekdayRank(day)
  if (Number.isNaN(start.getTime()) || rank === 99) return ''
  const startRank = ((start.getDay() + 6) % 7) + 1
  start.setDate(start.getDate() + ((rank - startRank + 7) % 7))
  return isoDate(start)
}

function arrivalKey(row: ProgrammeRow) {
  return `${row.day}::${row.id}`
}

function formatDate(value: string) {
  if (!value) return 'Not entered'
  const date = new Date(`${value}T12:00:00`)
  if (Number.isNaN(date.getTime())) return value
  return date.toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' })
}

function accommodationDetails(flatIds: string[] = []) {
  if (!flatIds.length) return { buildings: 'Not allocated', flats: 'Not allocated' }
  const byBuilding = new Map<number, number[]>()
  flatIds.forEach((flatId) => {
    const [building, flat] = flatId.split('-').map(Number)
    if (!building || !flat) return
    const values = byBuilding.get(building) ?? []
    values.push(flat)
    byBuilding.set(building, values)
  })
  const entries = [...byBuilding.entries()].sort(([a], [b]) => a - b)
  return {
    buildings: entries.map(([building]) => buildings[building - 1] ?? `Building ${building}`).join(', '),
    flats: entries.map(([building, flats]) => `${buildings[building - 1] ?? `Building ${building}`}: ${flats.sort((a, b) => a - b).join(', ')}`).join(' · '),
  }
}

function weatherLabel(code: number) {
  if (code === 0) return 'Clear'
  if ([1, 2].includes(code)) return 'Partly cloudy'
  if (code === 3) return 'Overcast'
  if ([45, 48].includes(code)) return 'Foggy'
  if ([51, 53, 55, 56, 57].includes(code)) return 'Drizzle'
  if ([61, 63, 65, 66, 67, 80, 81, 82].includes(code)) return 'Rain'
  if ([71, 73, 75, 77, 85, 86].includes(code)) return 'Snow'
  if ([95, 96, 99].includes(code)) return 'Thunderstorms'
  return 'Current conditions'
}

export default function ManagerDisplay() {
  const [state, setState] = useState<LiveState>({ programme: null, staff: [], arrivalAssignments: {} })
  const [dayOff, setDayOff] = useState<DayOffEntry[]>([])
  const [weather, setWeather] = useState<WeatherState | null>(null)
  const [now, setNow] = useState(new Date())

  async function loadLiveState() {
    const { data } = await supabase.from('app_live_state').select('state').eq('id', 'main').maybeSingle()
    if (data?.state) setState(data.state as LiveState)
  }

  async function loadDayOff() {
    const today = isoDate(new Date())
    const { data } = await supabase.from('staff_days_off').select('staff_id,staff_name,day,status').eq('day', today)
    setDayOff((data ?? []) as DayOffEntry[])
  }

  async function loadWeather() {
    try {
      const response = await fetch('https://api.open-meteo.com/v1/forecast?latitude=52.69&longitude=0.95&current=temperature_2m,weather_code,wind_speed_10m&timezone=Europe%2FLondon')
      if (!response.ok) throw new Error('Weather unavailable')
      const data = await response.json()
      setWeather({
        temperature: Math.round(Number(data.current.temperature_2m)),
        wind: Math.round(Number(data.current.wind_speed_10m)),
        code: Number(data.current.weather_code),
      })
    } catch {
      setWeather(null)
    }
  }

  useEffect(() => {
    void loadLiveState()
    void loadDayOff()
    void loadWeather()
    const timer = window.setInterval(() => {
      setNow(new Date())
      void loadLiveState()
      void loadDayOff()
      void loadWeather()
    }, 30000)
    const liveChannel = supabase.channel('manager-display-live').on('postgres_changes', { event: '*', schema: 'public', table: 'app_live_state', filter: 'id=eq.main' }, () => void loadLiveState()).subscribe()
    const absenceChannel = supabase.channel('manager-display-absence').on('postgres_changes', { event: '*', schema: 'public', table: 'staff_days_off' }, () => void loadDayOff()).subscribe()
    return () => {
      window.clearInterval(timer)
      void supabase.removeChannel(liveChannel)
      void supabase.removeChannel(absenceChannel)
    }
  }, [])

  const today = isoDate(now)
  const unavailableIds = useMemo(() => new Set(dayOff.filter((entry) => ['off', 'hol', 'sick'].includes(entry.status)).map((entry) => entry.staff_id)), [dayOff])
  const sickStaff = useMemo(() => dayOff.filter((entry) => entry.status === 'sick').map((entry) => entry.staff_name).sort(), [dayOff])
  const activeStaff = useMemo(() => state.staff.filter((member) => !unavailableIds.has(member.id)), [state.staff, unavailableIds])
  const loanStaff = useMemo(() => activeStaff.filter((member) => member.employmentType === 'loan' && (!member.startDate || member.startDate <= today) && (!member.loanEndDate || member.loanEndDate >= today)).map((member) => member.name).sort(), [activeStaff, today])

  const schools = useMemo(() => {
    const programme = state.programme
    if (!programme) return []
    return (programme.schoolDetails ?? [])
      .filter((school) => school.arrivalDate <= today && school.departureDate >= today)
      .map((school) => {
        const arrivalRow = programme.rows.find((row) => row.session === '3' && row.schoolLabel?.trim().toLowerCase() === school.schoolName.trim().toLowerCase())
          ?? programme.rows.find((row) => row.schoolLabel?.trim().toLowerCase() === school.schoolName.trim().toLowerCase())
        const assignment = arrivalRow ? state.arrivalAssignments[arrivalKey(arrivalRow)] : undefined
        const groups = school.groupNumbers?.length
          ? school.groupNumbers
          : [...new Set(programme.rows.flatMap((row) => row.schoolLabel?.trim().toLowerCase() === school.schoolName.trim().toLowerCase() ? row.cells.map((cell) => cell.group) : []))]
        const accommodation = accommodationDetails(assignment?.flatIds)
        return {
          id: school.id,
          name: school.schoolName,
          arrivalDate: school.arrivalDate,
          departureDate: school.departureDate,
          groupCount: groups.length,
          accommodation,
          notes: assignment?.notes?.trim() ?? '',
        }
      })
      .sort((a, b) => a.name.localeCompare(b.name))
  }, [state.programme, state.arrivalAssignments, today])

  return <main className="manager-display-shell" aria-label="Manager display">
    <header className="manager-display-header">
      <div className="manager-display-brand">
        <img src={`${import.meta.env.BASE_URL}manor-adventure-logo.png`} alt="Manor Adventure" />
        <div><p>NORFOLK LAKES</p><h1>Manager Display</h1></div>
      </div>
      <div className="manager-display-clock">
        <CalendarDays />
        <strong>{now.toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</strong>
        <span>{now.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })}</span>
      </div>
    </header>

    <section className="manager-display-metrics">
      <article><School /><div><span>Schools on site</span><strong>{schools.length}</strong></div></article>
      <article><Users /><div><span>Staff in today</span><strong>{activeStaff.length}</strong></div></article>
      <article className={sickStaff.length ? 'metric-alert' : ''}><HeartPulse /><div><span>Sick today</span><strong>{sickStaff.length}</strong><small>{sickStaff.length ? sickStaff.join(', ') : 'None'}</small></div></article>
      <article><Users /><div><span>Loan staff</span><strong>{loanStaff.length}</strong><small>{loanStaff.length ? loanStaff.join(', ') : 'None today'}</small></div></article>
      <article className="weather-metric"><CloudSun /><div><span>Norfolk Lakes weather</span><strong>{weather ? `${weather.temperature}°C` : '—'}</strong><small>{weather ? `${weatherLabel(weather.code)} · Wind ${weather.wind} km/h` : 'Weather unavailable'}</small></div></article>
    </section>

    <section className="manager-display-schools">
      <div className="manager-display-section-title"><Building2 /><h2>Schools currently on site</h2></div>
      {schools.length ? <div className={`manager-school-grid count-${Math.min(schools.length, 6)}`}>
        {schools.map((school) => <article className="manager-school-card" key={school.id}>
          <header><h3>{school.name}</h3><span>{school.groupCount} group{school.groupCount === 1 ? '' : 's'}</span></header>
          <div className="manager-school-details">
            <div><span>Arrived</span><strong>{formatDate(school.arrivalDate)}</strong></div>
            <div><span>Leaving</span><strong>{formatDate(school.departureDate)}</strong></div>
            <div><span>Building</span><strong>{school.accommodation.buildings}</strong></div>
            <div><span>Flats</span><strong>{school.accommodation.flats}</strong></div>
          </div>
          {school.notes && <section className="manager-school-notes"><span>School notes</span><p>{school.notes}</p></section>}
        </article>)}
      </div> : <div className="manager-display-empty"><School /><h3>No schools are currently on site</h3></div>}
    </section>

    <footer>Live read-only manager display · updates automatically every 30 seconds</footer>
  </main>
}
