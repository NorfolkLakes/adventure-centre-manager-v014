import { FormEvent, useEffect, useState } from 'react'
import type { Session } from '@supabase/supabase-js'
import { CalendarDays, LogOut, ShieldCheck } from 'lucide-react'
import ManagerApp from './ManagerApp'
import { supabase } from './lib/supabase'

type Profile = {
  id: string
  email: string
  display_name: string | null
  role: 'manager' | 'staff' | 'centreManager' | 'activityManager' | 'teamLeader'
}


const WEEKDAY_ORDER: Record<string, number> = {
  MON: 1, MONDAY: 1,
  TUE: 2, TUES: 2, TUESDAY: 2,
  WED: 3, WEDNESDAY: 3,
  THU: 4, THUR: 4, THURS: 4, THURSDAY: 4,
  FRI: 5, FRIDAY: 5,
  SAT: 6, SATURDAY: 6,
  SUN: 7, SUNDAY: 7,
}

function weekdayRank(value: string) {
  const key = value.trim().toUpperCase().replace(/[^A-Z]/g, '')
  return WEEKDAY_ORDER[key] ?? 99
}

function displayProgrammeDay(value: string) {
  const key = value.trim().toUpperCase().replace(/[^A-Z]/g, '')
  const names: Record<string, string> = {
    MON: 'Monday', MONDAY: 'Monday',
    TUE: 'Tuesday', TUES: 'Tuesday', TUESDAY: 'Tuesday',
    WED: 'Wednesday', WEDNESDAY: 'Wednesday',
    THU: 'Thursday', THUR: 'Thursday', THURS: 'Thursday', THURSDAY: 'Thursday',
    FRI: 'Friday', FRIDAY: 'Friday',
    SAT: 'Saturday', SATURDAY: 'Saturday',
    SUN: 'Sunday', SUNDAY: 'Sunday',
  }
  return names[key] ?? value
}

type RotaDuty = {
  id: string
  programme_name: string
  day: string
  session: string
  activity_code: string
  activity_name: string
  group_numbers: number[]
  duty_type: string
  staff_name: string
  school_name: string | null
  building_name: string | null
  party_leader_name: string | null
  arrival_time: string | null
  departure_day: string | null
  departure_time: string | null
}

function App() {
  const [session, setSession] = useState<Session | null>(null)
  const [profile, setProfile] = useState<Profile | null>(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session)
    })

    const { data: listener } = supabase.auth.onAuthStateChange(
      (_event, nextSession) => {
        setSession(nextSession)
      },
    )

    return () => listener.subscription.unsubscribe()
  }, [])

  useEffect(() => {
    async function loadProfile() {
      if (!session?.user) {
        setProfile(null)
        setLoading(false)
        return
      }

      setLoading(true)
      const { data, error } = await supabase
        .from('profiles')
        .select('id,email,display_name,role')
        .eq('id', session.user.id)
        .single()

      if (error) {
        setProfile({
          id: session.user.id,
          email: session.user.email ?? '',
          display_name: null,
          role: 'staff',
        })
      } else {
        setProfile(data as Profile)
      }
      setLoading(false)
    }

    loadProfile()
  }, [session])

  async function signOut() {
    await supabase.auth.signOut()
  }

  if (loading) {
    return <div className="auth-loading">Loading Adventure Centre Manager…</div>
  }

  if (!session) {
    return <LoginScreen />
  }

  if (profile?.role === 'manager' || profile?.role === 'centreManager' || profile?.role === 'activityManager' || profile?.role === 'teamLeader') {
    return (
      <ManagerApp
        accountEmail={profile.email || session.user.email || ''}
        displayName={profile.display_name ?? null}
        onSignOut={signOut}
        accountRole={profile.role === 'activityManager' ? 'activityManager' : profile.role === 'teamLeader' ? 'teamLeader' : 'centreManager'}
      />
    )
  }


  return (
    <StaffRota
      accountEmail={profile?.email || session.user.email || ''}
      displayName={profile?.display_name ?? null}
      onSignOut={signOut}
    />
  )
}

function TeamLeaderHolidayView({ accountEmail, onSignOut }: { accountEmail: string; onSignOut: () => void }) {
  const [holidays, setHolidays] = useState<{id:string;staff_name:string;start_date:string;end_date:string;note:string|null}[]>([])
  const [month, setMonth] = useState(() => new Date(new Date().getFullYear(), new Date().getMonth(), 1))
  useEffect(() => {
    async function load() {
      const { data } = await supabase.from('staff_holidays').select('id,staff_name,start_date,end_date,note').order('start_date')
      setHolidays((data ?? []) as typeof holidays)
    }
    load()
    const channel = supabase.channel('team-leader-holidays').on('postgres_changes', { event: '*', schema: 'public', table: 'staff_holidays' }, load).subscribe()
    return () => { supabase.removeChannel(channel) }
  }, [])
  const dateKey = (date: Date) => `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`
  const first = new Date(month.getFullYear(), month.getMonth(), 1); const start = new Date(first); start.setDate(first.getDate()-((first.getDay()+6)%7))
  const days = Array.from({length:42},(_,i)=>{const d=new Date(start);d.setDate(start.getDate()+i);return d})
  return <main className="app-shell team-leader-shell">
    <header className="app-header"><div><p className="eyebrow">Norfolk Lakes</p><h1>Holiday calendar</h1><small>{accountEmail} · Team Leader · View only</small></div><button className="secondary-action" onClick={onSignOut}><LogOut size={17}/>Sign out</button></header>
    <section className="panel"><div className="holiday-summary"><h2>{month.toLocaleDateString('en-GB',{month:'long',year:'numeric'})}</h2><div className="holiday-month-actions"><button className="secondary-action" onClick={()=>setMonth(new Date(month.getFullYear(),month.getMonth()-1,1))}>Previous</button><button className="secondary-action" onClick={()=>setMonth(new Date())}>Today</button><button className="secondary-action" onClick={()=>setMonth(new Date(month.getFullYear(),month.getMonth()+1,1))}>Next</button></div></div>
    <div className="holiday-weekdays">{['Mon','Tue','Wed','Thu','Fri','Sat','Sun'].map(day=><strong key={day}>{day}</strong>)}</div><div className="holiday-calendar">{days.map(date=>{const key=dateKey(date); const entries=holidays.filter(h=>h.start_date<=key&&h.end_date>=key);return <article key={key} className={`holiday-day ${date.getMonth()!==month.getMonth()?'outside':''} ${key===dateKey(new Date())?'today':''}`}><span className="holiday-date">{date.getDate()}</span>{entries.map(h=><div className="holiday-entry readonly" key={h.id}>{h.staff_name}</div>)}</article>})}</div></section>
  </main>
}

function LoginScreen() {
  const [mode, setMode] = useState<'login' | 'signup'>('login')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [message, setMessage] = useState('')
  const [selectedDay, setSelectedDay] = useState('')
  const [availability, setAvailability] = useState<Record<string, 'available' | 'holiday' | 'sick'>>({})
  const [busy, setBusy] = useState(false)

  async function submit(event: FormEvent) {
    event.preventDefault()
    setBusy(true)
    setMessage('')

    if (mode === 'login') {
      const { error } = await supabase.auth.signInWithPassword({
        email,
        password,
      })
      setMessage(error ? error.message : 'Signed in.')
    } else {
      const { error } = await supabase.auth.signUp({
        email,
        password,
      })
      setMessage(
        error
          ? error.message
          : 'Account created. Check your email if confirmation is enabled.',
      )
    }

    setBusy(false)
  }

  return (
    <main className="login-page">
      <section className="login-card">
        <div className="login-brand">
          <ShieldCheck size={34} />
          <div>
            <p className="eyebrow">Norfolk Lakes</p>
            <h1>Adventure Centre Manager</h1>
          </div>
        </div>

        <h2>{mode === 'login' ? 'Sign in' : 'Create staff account'}</h2>
        <p>
          Managers can publish the rota. Staff see only duties assigned to
          their login email.
        </p>

        <form onSubmit={submit}>
          <label>Email address</label>
          <input
            type="email"
            required
            value={email}
            onChange={(event) => setEmail(event.target.value)}
          />

          <label>Password</label>
          <input
            type="password"
            minLength={6}
            required
            value={password}
            onChange={(event) => setPassword(event.target.value)}
          />

          <button className="primary login-submit" disabled={busy}>
            {busy
              ? 'Please wait…'
              : mode === 'login'
                ? 'Sign in'
                : 'Create account'}
          </button>
        </form>

        {message && <div className="auth-message">{message}</div>}

        <button
          className="auth-switch"
          onClick={() => {
            setMode(mode === 'login' ? 'signup' : 'login')
            setMessage('')
          }}
        >
          {mode === 'login'
            ? 'New staff member? Create an account'
            : 'Already have an account? Sign in'}
        </button>
      </section>
    </main>
  )
}

function StaffRota({
  accountEmail,
  displayName,
  onSignOut,
}: {
  accountEmail: string
  displayName: string | null
  onSignOut: () => void
}) {
  const [duties, setDuties] = useState<RotaDuty[]>([])
  const [loading, setLoading] = useState(true)
  const [message, setMessage] = useState('')
  const [selectedDay, setSelectedDay] = useState('')
  const [holidays, setHolidays] = useState<{ id: string; start_date: string; end_date: string; note: string | null }[]>([])
  const [holidayMonth, setHolidayMonth] = useState(() => new Date())

  async function loadMyHolidays() {
    const monthStart = new Date(holidayMonth.getFullYear(), holidayMonth.getMonth(), 1)
    const monthEnd = new Date(holidayMonth.getFullYear(), holidayMonth.getMonth() + 1, 0)
    const dateKey = (date: Date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`
    const { data, error } = await supabase
      .from('staff_holidays')
      .select('id,start_date,end_date,note')
      .eq('staff_email', accountEmail.toLowerCase())
      .lte('start_date', dateKey(monthEnd))
      .gte('end_date', dateKey(monthStart))
      .order('start_date')
    if (error) setMessage(error.message)
    else setHolidays((data ?? []) as typeof holidays)
  }

  async function loadDuties() {
    setLoading(true)
    const { data, error } = await supabase
      .from('rota_assignments')
      .select(
        'id,programme_name,day,session,activity_code,activity_name,group_numbers,duty_type,staff_name,school_name,building_name,party_leader_name,arrival_time,departure_day,departure_time',
      )
      .order('day')
      .order('session')

    if (error) {
      setMessage(error.message)
    } else {
      const nextDuties = ((data ?? []) as RotaDuty[]).sort((a, b) => {
        const dayDifference = weekdayRank(a.day) - weekdayRank(b.day)
        if (dayDifference !== 0) return dayDifference
        return Number(a.session) - Number(b.session)
      })
      setDuties(nextDuties)
      setSelectedDay((current) =>
        current && nextDuties.some((duty) => duty.day === current)
          ? current
          : nextDuties[0]?.day || '',
      )
    }
    setLoading(false)
  }

  useEffect(() => {
    loadDuties()

    const channel = supabase
      .channel('staff-rota-updates')
      .on(
        'postgres_changes',
        { event: '*', schema: 'public', table: 'rota_assignments' },
        () => loadDuties(),
      )
      .subscribe()

    const holidayChannel = supabase
      .channel(`my-holidays-${accountEmail}`)
      .on(
        'postgres_changes',
        { event: '*', schema: 'public', table: 'staff_holidays' },
        () => loadMyHolidays(),
      )
      .subscribe()

    return () => {
      supabase.removeChannel(channel)
      supabase.removeChannel(holidayChannel)
    }
  }, [])

  useEffect(() => {
    loadMyHolidays()
  }, [holidayMonth, accountEmail])

  const days = Array.from(new Set(duties.map((duty) => duty.day))).sort((a, b) => weekdayRank(a) - weekdayRank(b))

  return (
    <div className="staff-app">
      <header className="staff-header">
        <div>
          <p className="eyebrow">My rota</p>
          <h1>{displayName || duties[0]?.staff_name || accountEmail}</h1>
          <small>{accountEmail}</small>
        </div>
        <button onClick={onSignOut}>
          <LogOut size={17} />
          Sign out
        </button>
      </header>

      <main className="staff-content">
        {loading && <div className="staff-empty">Loading your duties…</div>}
        {message && <div className="auth-message">{message}</div>}

        {!loading && !duties.length && (
          <section className="staff-empty">
            <CalendarDays size={40} />
            <h2>No duties published yet</h2>
            <p>
              Ask your manager to add your login email to your staff record
              and make sure your login email matches your staff record. Rota changes appear automatically.
            </p>
          </section>
        )}

        {days.length > 0 && (
          <>
            <div className="staff-day-tabs">
              {days.map((day) => (
                <button
                  key={day}
                  className={selectedDay === day ? 'active' : ''}
                  onClick={() => setSelectedDay(day)}
                >
                  {displayProgrammeDay(day)}
                </button>
              ))}
            </div>
            <section className="staff-day">
              <div className="staff-day-heading">
                <div>
                  <p className="eyebrow">Your duties</p>
                  <h2>{displayProgrammeDay(selectedDay)}</h2>
                </div>
                <span>
                  {duties.filter((duty) => duty.day === selectedDay).length}{' '}
                  duties
                </span>
              </div>
              <div className="my-rota-list">
                {duties
                  .filter((duty) => duty.day === selectedDay)
                  .map((duty) => (
                    <article
                      className={`my-rota-card duty-${duty.duty_type}`}
                      key={duty.id}
                    >
                      <div className="my-rota-card-top">
                        <div className="my-rota-session">
                          Session {duty.session}
                        </div>
                        <span className="live-badge">Live</span>
                      </div>
                      <h3>{duty.duty_type.startsWith('arrival_') ? 'Arrivals' : duty.activity_name}</h3>
                      {duty.school_name && (
                        <p className="school-name">{duty.school_name}</p>
                      )}
                      {duty.duty_type.startsWith('arrival_') && (
                        <div className="arrival-duty-details">
                          {duty.school_name && <p><strong>School:</strong> {duty.school_name}</p>}
                          {duty.group_numbers.length > 0 && <p><strong>Group{duty.group_numbers.length > 1 ? 's' : ''}:</strong> {duty.group_numbers.map((group) => `G${group}`).join(', ')}</p>}
                          {duty.building_name && <p><strong>Building:</strong> {duty.building_name}</p>}
                          {duty.party_leader_name && duty.duty_type !== 'arrival_leader' && <p><strong>Party Leader:</strong> {duty.party_leader_name}</p>}
                        </div>
                      )}
                      <div className="group-pill">
                        {duty.group_numbers.length
                          ? `Group${
                              duty.group_numbers.length > 1 ? 's' : ''
                            } ${duty.group_numbers.map((group) => `G${group}`).join(', ')}`
                          : 'No group'}
                      </div>
                    </article>
                  ))}
              </div>
            </section>
          </>
        )}

        <section className="my-holiday-panel">
          <div className="my-holiday-heading">
            <div>
              <p className="eyebrow">My holiday</p>
              <h2>{holidayMonth.toLocaleDateString('en-GB', { month: 'long', year: 'numeric' })}</h2>
            </div>
            <div className="my-holiday-actions">
              <button onClick={() => setHolidayMonth(new Date(holidayMonth.getFullYear(), holidayMonth.getMonth() - 1, 1))}>Previous</button>
              <button onClick={() => setHolidayMonth(new Date())}>This month</button>
              <button onClick={() => setHolidayMonth(new Date(holidayMonth.getFullYear(), holidayMonth.getMonth() + 1, 1))}>Next</button>
            </div>
          </div>
          {holidays.length ? (
            <div className="my-holiday-list">
              {holidays.map((holiday) => (
                <article key={holiday.id}>
                  <strong>{new Date(`${holiday.start_date}T12:00:00`).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })} – {new Date(`${holiday.end_date}T12:00:00`).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })}</strong>
                  {holiday.note && <span>{holiday.note}</span>}
                </article>
              ))}
            </div>
          ) : (
            <p className="my-holiday-empty">No holiday booked this month.</p>
          )}
        </section>
      </main>
    </div>
  )
}

export default App
