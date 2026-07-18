import { useEffect, useMemo, useState } from 'react'
import { CalendarDays, Clock3, Maximize2, RefreshCw, Users } from 'lucide-react'
import { supabase } from './lib/supabase'
import type { ProgrammeImport, StaffMember, StaffingAssignment } from './types'

type DayOffStatus = 'off' | 'hol' | 'sick' | 'am_off' | 'pm_off'
type StaffDayOff = { id:string; staff_id:string; staff_email:string; staff_name:string; day:string; status:DayOffStatus; note:string|null }
type LiveState = { programme: ProgrammeImport | null; staff: StaffMember[]; activities: {code:string;name:string}[]; assignments: StaffingAssignment }

const sessionOrder = ['1','2','3','4','5']
const weekdayNames = ['SUN','MON','TUE','WED','THU','FRI','SAT']

function cellKey(rowId: string, group: number) { return `${rowId}::${group}` }
function isoDate(date: Date) { return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}` }
function weekdayRank(value: string) {
  const key = value.toUpperCase().replace(/[^A-Z]/g,'')
  return ({MON:1,MONDAY:1,TUE:2,TUES:2,TUESDAY:2,WED:3,WEDNESDAY:3,THU:4,THUR:4,THURS:4,THURSDAY:4,FRI:5,FRIDAY:5,SAT:6,SATURDAY:6,SUN:7,SUNDAY:7} as Record<string,number>)[key] ?? 99
}
function dateForProgrammeDay(programme: ProgrammeImport, day: string) {
  if (/^\d{4}-\d{2}-\d{2}$/.test(day)) return day
  if (programme.startDate) {
    const start = new Date(`${programme.startDate}T12:00:00`)
    const rank = weekdayRank(day)
    if (!Number.isNaN(start.getTime()) && rank < 99) {
      const startRank = ((start.getDay()+6)%7)+1
      const result = new Date(start)
      result.setDate(start.getDate()+((rank-startRank+7)%7))
      return isoDate(result)
    }
  }
  return ''
}
function statusLabel(status: DayOffStatus) {
  return ({off:'OFF',hol:'HOLIDAY',sick:'SICK',am_off:'AM OFF',pm_off:'PM OFF'} as Record<DayOffStatus,string>)[status]
}

export default function StaffRoomDisplay({ onExit }: { onExit: () => void }) {
  const [state, setState] = useState<LiveState>({ programme:null, staff:[], activities:[], assignments:{} })
  const [daysOff, setDaysOff] = useState<StaffDayOff[]>([])
  const [now, setNow] = useState(new Date())
  const [updatedAt, setUpdatedAt] = useState('')

  async function load() {
    const [{ data: live }, { data: off }] = await Promise.all([
      supabase.from('app_live_state').select('state,updated_at').eq('id','main').maybeSingle(),
      supabase.from('staff_days_off').select('id,staff_id,staff_email,staff_name,day,status,note').gte('day', isoDate(new Date())).lte('day', isoDate(new Date(Date.now()+7*86400000))),
    ])
    if (live?.state) setState(live.state as LiveState)
    if (live?.updated_at) setUpdatedAt(live.updated_at)
    setDaysOff((off ?? []) as StaffDayOff[])
  }

  useEffect(() => {
    load()
    const timer = window.setInterval(() => { setNow(new Date()); load() }, 30000)
    const channel = supabase.channel('staff-room-display-live')
      .on('postgres_changes',{event:'*',schema:'public',table:'app_live_state',filter:'id=eq.main'},load)
      .on('postgres_changes',{event:'*',schema:'public',table:'staff_days_off'},load)
      .subscribe()
    return () => { window.clearInterval(timer); supabase.removeChannel(channel) }
  }, [])

  const todayIso = isoDate(now)
  const programmeDay = useMemo(() => state.programme?.rows.find(row => dateForProgrammeDay(state.programme!, row.day) === todayIso)?.day ?? '', [state.programme, todayIso])
  const activityMap = useMemo(() => new Map(state.activities.map(item => [item.code,item.name])), [state.activities])
  const duties = useMemo(() => {
    const result = new Map<string, Record<string,string[]>>()
    if (!state.programme || !programmeDay) return result
    for (const row of state.programme.rows.filter(item => item.day === programmeDay)) {
      for (const cell of row.cells) {
        const staffId = state.assignments[cellKey(row.id, cell.group)]
        if (!staffId) continue
        const sessions = result.get(staffId) ?? {}
        const label = `${cell.activityCode}${cell.group ? ` G${cell.group}` : ''}`
        sessions[row.session] = [...(sessions[row.session] ?? []), label]
        result.set(staffId, sessions)
      }
    }
    return result
  }, [state.programme, state.assignments, programmeDay])

  const todayDaysOff = useMemo(() => daysOff.filter(item => item.day === todayIso), [daysOff, todayIso])
  const upcomingHolidays = useMemo(() => daysOff
    .filter(item => item.day > todayIso && item.status === 'hol')
    .sort((a,b) => a.day.localeCompare(b.day) || a.staff_name.localeCompare(b.staff_name)), [daysOff, todayIso])
  const offByStaff = useMemo(() => new Map(todayDaysOff.map(item => [item.staff_id,item])), [todayDaysOff])
  const visibleStaff = useMemo(() => {
    const roleRank = (member: StaffMember) => member.role === 'centreManager' ? 0 : member.role === 'activityManager' ? 1 : member.role === 'teamLeader' || member.teamLeader ? 2 : 3
    return [...state.staff].sort((a,b) => roleRank(a) - roleRank(b) || a.name.localeCompare(b.name))
  }, [state.staff])
  const offGroups = useMemo(() => todayDaysOff.reduce<Record<string,StaffDayOff[]>>((acc,item) => { (acc[item.status] ??= []).push(item); return acc },{}), [todayDaysOff])

  async function fullScreen() { try { await document.documentElement.requestFullscreen() } catch { /* browser may block */ } }

  return <main className="staff-display-shell">
    <header className="staff-display-header">
      <div className="staff-display-brand"><img src={`${import.meta.env.BASE_URL}manor-adventure-logo.png`} alt="Manor Adventure"/><div><p>NORFOLK LAKES</p><h1>Daily Staffing</h1></div></div>
      <div className="staff-display-date"><CalendarDays/><strong>{now.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long',year:'numeric'})}</strong><span><Clock3 size={18}/>{now.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'})}</span></div>
      <div className="staff-display-actions"><button onClick={() => void load()}><RefreshCw/>Refresh</button><button onClick={fullScreen}><Maximize2/>Full screen</button><button onClick={onExit}>Exit display</button></div>
    </header>

    {!state.programme || !programmeDay ? <section className="staff-display-empty"><Users size={42}/><h2>No programme is live for today</h2><p>The display will update automatically when a programme and staffing are published.</p></section> : <>
      <section className="staff-display-summary"><article><span>Programme</span><strong>{state.programme.title}</strong></article><article><span>Staff shown</span><strong>{visibleStaff.length}</strong></article><article><span>Off today</span><strong>{todayDaysOff.length}</strong></article><article><span>Last updated</span><strong>{updatedAt ? new Date(updatedAt).toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'}) : '—'}</strong></article></section>
      <div className="staff-display-layout">
        <section className="staff-display-rota"><table><thead><tr><th>Staff member</th>{sessionOrder.map(s => <th key={s}>Session {s}</th>)}</tr></thead><tbody>{visibleStaff.map(member => { const off = offByStaff.get(member.id); return <tr key={member.id} className={off ? `display-off display-${off.status}` : ''}><th>{member.name}<small>{off ? statusLabel(off.status) : member.role === 'centreManager' ? 'Head of Centre' : member.role === 'activityManager' ? 'Activities Manager' : member.role === 'teamLeader' ? 'Team Leader' : 'Instructor'}</small></th>{sessionOrder.map(session => { const partialOff = off?.status === 'am_off' && ['1','2'].includes(session) || off?.status === 'pm_off' && ['3','4','5'].includes(session); const fullOff = off && ['off','hol','sick'].includes(off.status); const values = duties.get(member.id)?.[session] ?? []; return <td key={session} className={fullOff || partialOff ? 'status-cell' : values.length ? 'duty-cell' : 'empty-cell'}>{fullOff || partialOff ? statusLabel(off!.status) : values.length ? values.map(value => <span key={value} title={activityMap.get(value.split(' ')[0])}>{value}</span>) : '—'}</td>})}</tr>})}</tbody></table></section>
        <aside className="staff-display-off-panel"><h2>Days off</h2>{(['hol','sick','off','am_off','pm_off'] as DayOffStatus[]).map(status => <section key={status} className={`off-list off-${status}`}><h3>{status === 'hol' ? 'HOLIDAY TODAY' : statusLabel(status)}</h3>{offGroups[status]?.length ? offGroups[status].map(item => <div key={item.id}><strong>{item.staff_name}</strong>{item.note && <small>{item.note}</small>}</div>) : <p>None</p>}</section>)}<section className="off-list upcoming-holiday"><h3>UPCOMING HOLIDAYS · NEXT 7 DAYS</h3>{upcomingHolidays.length ? upcomingHolidays.map(item => <div key={`${item.id}-${item.day}`}><strong>{new Date(`${item.day}T12:00:00`).toLocaleDateString('en-GB',{weekday:'short',day:'numeric',month:'short'})} — {item.staff_name}</strong>{item.note && <small>{item.note}</small>}</div>) : <p>None</p>}</section></aside>
      </div>
    </>}
    <footer className="staff-display-footer">Live read-only display · refreshes automatically every 30 seconds</footer>
  </main>
}
