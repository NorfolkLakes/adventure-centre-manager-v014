import { FormEvent, useEffect, useState } from 'react'
import type { Session } from '@supabase/supabase-js'
import { CalendarDays, LogOut, ShieldCheck } from 'lucide-react'
import ManagerApp from './ManagerApp'
import { supabase } from './lib/supabase'

type Profile = {
  id: string
  email: string
  display_name: string | null
  role: 'manager' | 'staff'
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

  if (profile?.role === 'manager') {
    return (
      <ManagerApp
        accountEmail={profile.email || session.user.email || ''}
        onSignOut={signOut}
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

function LoginScreen() {
  const [mode, setMode] = useState<'login' | 'signup'>('login')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [message, setMessage] = useState('')
  const [selectedDay, setSelectedDay] = useState('')
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

  async function loadDuties() {
    setLoading(true)
    const { data, error } = await supabase
      .from('rota_assignments')
      .select(
        'id,programme_name,day,session,activity_code,activity_name,group_numbers,duty_type,staff_name,school_name',
      )
      .order('day')
      .order('session')

    if (error) {
      setMessage(error.message)
    } else {
      const nextDuties = (data ?? []) as RotaDuty[]
      setDuties(nextDuties)
      setSelectedDay((current) => current || nextDuties[0]?.day || '')
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

    return () => {
      supabase.removeChannel(channel)
    }
  }, [])

  const days = Array.from(new Set(duties.map((duty) => duty.day)))

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
              and press Publish staff rota.
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
                  {day}
                </button>
              ))}
            </div>
            <section className="staff-day">
              <div className="staff-day-heading">
                <div>
                  <p className="eyebrow">Your duties</p>
                  <h2>{selectedDay}</h2>
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
                      <h3>{duty.activity_name}</h3>
                      {duty.school_name && (
                        <p className="school-name">{duty.school_name}</p>
                      )}
                      <div className="group-pill">
                        {duty.group_numbers.length
                          ? `Group${
                              duty.group_numbers.length > 1 ? 's' : ''
                            } ${duty.group_numbers.join(', ')}`
                          : 'No group'}
                      </div>
                    </article>
                  ))}
              </div>
            </section>
          </>
        )}
      </main>
    </div>
  )
}

export default App
