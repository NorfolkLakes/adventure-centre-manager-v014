import { useEffect, useMemo, useState } from 'react'
import { AlertTriangle, CheckCircle2, CloudLightning, RadioTower } from 'lucide-react'

type LightningStrike = {
  distance_mi?: number
  distance_km?: number
  bearing_cardinal?: string
  past_mins?: number
  timestamp_utc?: string
}

type LightningResponse = {
  lightning?: LightningStrike[]
}

export type LightningState = {
  status: 'safe' | 'danger' | 'unavailable'
  nearestMiles: number | null
  direction: string
  lastStrikeMinutes: number | null
  restartAt: Date | null
  updatedAt: Date | null
  message: string
}

const LATITUDE = 52.69
const LONGITUDE = 0.95
const TRIGGER_MILES = 7
const SEARCH_DISTANCE_KM = 25
const CLEAR_MINUTES = 30
const POLL_MS = 5 * 60 * 1000

function toMiles(strike: LightningStrike) {
  if (Number.isFinite(Number(strike.distance_mi))) return Number(strike.distance_mi)
  if (Number.isFinite(Number(strike.distance_km))) return Number(strike.distance_km) * 0.621371
  return Number.POSITIVE_INFINITY
}

async function fetchLightning(): Promise<LightningResponse> {
  const configuredProxy = String(import.meta.env.VITE_LIGHTNING_PROXY_URL ?? '').trim()
  const supabaseUrl = String(import.meta.env.VITE_SUPABASE_URL ?? '').trim().replace(/\/$/, '')
  const proxyUrl = configuredProxy || (supabaseUrl ? `${supabaseUrl}/functions/v1/lightning` : '')
  const apiKey = String(import.meta.env.VITE_WEATHERBIT_API_KEY ?? '').trim()

  if (proxyUrl) {
    const separator = proxyUrl.includes('?') ? '&' : '?'
    const response = await fetch(`${proxyUrl}${separator}lat=${LATITUDE}&lon=${LONGITUDE}&search_distance_km=${SEARCH_DISTANCE_KM}&search_mins=45&sort=distance`)
    if (!response.ok) throw new Error(`Lightning service returned ${response.status}`)
    return response.json() as Promise<LightningResponse>
  }

  if (apiKey) {
    const params = new URLSearchParams({
      lat: String(LATITUDE),
      lon: String(LONGITUDE),
      search_distance_km: String(SEARCH_DISTANCE_KM),
      search_mins: '45',
      limit: '50',
      sort: 'distance',
      key: apiKey,
    })
    const response = await fetch(`https://api.weatherbit.io/v2.0/current/lightning?${params.toString()}`)
    if (response.status === 204) return { lightning: [] }
    if (!response.ok) throw new Error(`Lightning service returned ${response.status}`)
    return response.json() as Promise<LightningResponse>
  }

  throw new Error('Lightning API not configured')
}

export function useLightningMonitor() {
  const [state, setState] = useState<LightningState>({
    status: 'unavailable', nearestMiles: null, direction: '', lastStrikeMinutes: null,
    restartAt: null, updatedAt: null, message: 'Live lightning service is not configured.',
  })

  useEffect(() => {
    let cancelled = false

    async function load() {
      try {
        const data = await fetchLightning()
        if (cancelled) return
        const strikes = (data.lightning ?? [])
          .map((strike) => ({ ...strike, miles: toMiles(strike), minutes: Number(strike.past_mins ?? Number.POSITIVE_INFINITY) }))
          .filter((strike) => Number.isFinite(strike.miles) && Number.isFinite(strike.minutes))
          .sort((a, b) => a.miles - b.miles)
        const nearest = strikes[0]
        const nearbyRecent = strikes
          .filter((strike) => strike.miles <= TRIGGER_MILES && strike.minutes <= CLEAR_MINUTES)
          .sort((a, b) => a.minutes - b.minutes)[0]
        const lastNearby = strikes
          .filter((strike) => strike.miles <= TRIGGER_MILES)
          .sort((a, b) => a.minutes - b.minutes)[0]
        const now = new Date()

        if (nearbyRecent) {
          const restartAt = new Date(now.getTime() + Math.max(0, CLEAR_MINUTES - nearbyRecent.minutes) * 60_000)
          setState({
            status: 'danger',
            nearestMiles: nearbyRecent.miles,
            direction: nearbyRecent.bearing_cardinal ?? '',
            lastStrikeMinutes: nearbyRecent.minutes,
            restartAt,
            updatedAt: now,
            message: 'LIGHTNING WITHIN 7 MILES — CANCEL AND SUSPEND OUTDOOR SESSIONS.',
          })
          return
        }

        setState({
          status: 'safe',
          nearestMiles: nearest ? nearest.miles : null,
          direction: nearest?.bearing_cardinal ?? '',
          lastStrikeMinutes: lastNearby?.minutes ?? null,
          restartAt: null,
          updatedAt: now,
          message: 'No lightning detected within 7 miles during the last 30 minutes.',
        })
      } catch (error) {
        if (cancelled) return
        const message = error instanceof Error ? error.message : 'Lightning data unavailable'
        setState((previous) => ({ ...previous, status: 'unavailable', updatedAt: new Date(), message }))
      }
    }

    void load()
    const timer = window.setInterval(() => void load(), POLL_MS)
    return () => { cancelled = true; window.clearInterval(timer) }
  }, [])

  return state
}

function countdown(restartAt: Date | null, now: Date) {
  if (!restartAt) return ''
  const seconds = Math.max(0, Math.ceil((restartAt.getTime() - now.getTime()) / 1000))
  const minutes = Math.floor(seconds / 60)
  const remainder = seconds % 60
  return `${minutes}:${String(remainder).padStart(2, '0')}`
}

export function LightningStatusCard({ compact = false }: { compact?: boolean }) {
  const state = useLightningMonitor()
  const [now, setNow] = useState(new Date())
  useEffect(() => {
    const timer = window.setInterval(() => setNow(new Date()), 1000)
    return () => window.clearInterval(timer)
  }, [])

  const remaining = useMemo(() => countdown(state.restartAt, now), [state.restartAt, now])
  const className = `lightning-status-card lightning-${state.status}${compact ? ' lightning-compact' : ''}`
  const Icon = state.status === 'danger' ? AlertTriangle : state.status === 'safe' ? CheckCircle2 : RadioTower

  return <article className={className} aria-live="polite">
    <div className="lightning-icon"><Icon /></div>
    <div className="lightning-copy">
      <p className="eyebrow"><CloudLightning size={16}/> Lightning safety monitor</p>
      <h3>{state.status === 'danger' ? 'SESSIONS CANCELLED' : state.status === 'safe' ? '7-mile zone clear' : 'Lightning data unavailable'}</h3>
      <p>{state.message}</p>
      <div className="lightning-meta">
        {state.nearestMiles !== null && <span>Nearest: <strong>{state.nearestMiles.toFixed(1)} miles {state.direction}</strong></span>}
        {state.lastStrikeMinutes !== null && <span>Last nearby strike: <strong>{Math.round(state.lastStrikeMinutes)} min ago</strong></span>}
        {state.status === 'danger' && remaining && <span>Earliest review/restart: <strong>{remaining}</strong></span>}
        {state.updatedAt && <span>Updated: <strong>{state.updatedAt.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })}</strong></span>}
      </div>
    </div>
  </article>
}
