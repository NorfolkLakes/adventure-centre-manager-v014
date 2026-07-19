import { useEffect, useMemo, useState } from 'react'
import {
  ActivityIndicator,
  Alert,
  Pressable,
  SafeAreaView,
  ScrollView,
  StyleSheet,
  Text,
  TextInput,
  View,
} from 'react-native'
import { StatusBar } from 'expo-status-bar'
import type { Session } from '@supabase/supabase-js'
import { supabase } from './src/supabase'

type Duty = {
  id: string
  programme_name: string
  day: string
  session: string
  activity_name: string
  group_numbers: number[] | null
  duty_type: string
  staff_email: string
  staff_name: string
  school_name: string | null
  building_name: string | null
}

type Availability = 'available' | 'holiday' | 'sick'
type Tab = 'today' | 'timeOff' | 'notifications' | 'profile'

const weekdayOrder: Record<string, number> = {
  monday: 1, tuesday: 2, wednesday: 3, thursday: 4, friday: 5, saturday: 6, sunday: 7,
}

function sessionNumber(value: string) {
  const match = String(value).match(/[0-9]+(?:\.[0-9]+)?/)
  return match?.[0] ?? value
}

export default function App() {
  const [session, setSession] = useState<Session | null>(null)
  const [loading, setLoading] = useState(true)
  const [tab, setTab] = useState<Tab>('today')
  const [duties, setDuties] = useState<Duty[]>([])
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [busy, setBusy] = useState(false)
  const [selectedDay, setSelectedDay] = useState('')
  const [availability, setAvailability] = useState<Record<string, Availability>>({})

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session)
      setLoading(false)
    })
    const { data } = supabase.auth.onAuthStateChange((_event, next) => setSession(next))
    return () => data.subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (!session?.user.email) return
    const loginEmail = session.user.email.trim().toLowerCase()

    async function loadDuties() {
      const { data, error } = await supabase
        .from('rota_assignments')
        .select('id,programme_name,day,session,activity_name,group_numbers,duty_type,staff_email,staff_name,school_name,building_name')
        .eq('staff_email', loginEmail)
        .order('day')
        .order('session')
      if (error) Alert.alert('Rota unavailable', error.message)
      else setDuties((data ?? []) as Duty[])
    }

    async function loadAvailability() {
      const { data } = await supabase
        .from('staff_availability')
        .select('day,status')
        .eq('staff_email', loginEmail)
      const next: Record<string, Availability> = {}
      for (const row of (data ?? []) as { day: string; status: Availability }[]) next[row.day] = row.status
      setAvailability(next)
    }

    void loadDuties()
    void loadAvailability()
    const channel = supabase
      .channel(`staff-mobile-${session.user.id}`)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'rota_assignments' }, loadDuties)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'staff_availability' }, loadAvailability)
      .subscribe()
    return () => { void supabase.removeChannel(channel) }
  }, [session?.user.id, session?.user.email])

  const dutyDays = useMemo(() => Array.from(new Set(duties.map((duty) => duty.day))), [duties])
  const currentDay = useMemo(() => {
    const todayName = new Date().toLocaleDateString('en-GB', { weekday: 'long' }).toLowerCase()
    return dutyDays.find((day) => day.toLowerCase() === todayName) ?? dutyDays[0] ?? ''
  }, [dutyDays])
  const shownDay = selectedDay || currentDay
  const shownDuties = duties
    .filter((duty) => duty.day === shownDay)
    .sort((a, b) => Number(sessionNumber(a.session)) - Number(sessionNumber(b.session)))

  async function signIn() {
    setBusy(true)
    const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password })
    setBusy(false)
    if (error) Alert.alert('Sign in failed', error.message)
  }

  async function setDayStatus(day: string, status: Availability) {
    if (!session?.user.email) return
    setAvailability((current) => ({ ...current, [day]: status }))
    const { error } = await supabase.from('staff_availability').upsert({
      staff_email: session.user.email.trim().toLowerCase(), day, status,
    }, { onConflict: 'staff_email,day' })
    if (error) Alert.alert('Could not save', error.message)
  }

  if (loading) return <SafeAreaView style={styles.center}><ActivityIndicator size="large" /></SafeAreaView>
  if (!session) return (
    <SafeAreaView style={styles.login}>
      <StatusBar style="light" />
      <Text style={styles.kicker}>NORFOLK LAKES</Text>
      <Text style={styles.loginTitle}>Adventure Centre Staff</Text>
      <Text style={styles.loginCopy}>Sign in to see your live duties and record time off.</Text>
      <TextInput autoCapitalize="none" keyboardType="email-address" placeholder="Email" placeholderTextColor="#77908a" value={email} onChangeText={setEmail} style={styles.input} />
      <TextInput secureTextEntry placeholder="Password" placeholderTextColor="#77908a" value={password} onChangeText={setPassword} style={styles.input} />
      <Pressable onPress={signIn} disabled={busy} style={styles.primary}><Text style={styles.primaryText}>{busy ? 'Signing in…' : 'Sign in'}</Text></Pressable>
    </SafeAreaView>
  )

  return (
    <SafeAreaView style={styles.shell}>
      <StatusBar style="light" />
      <View style={styles.header}>
        <Text style={styles.kicker}>ADVENTURE CENTRE STAFF</Text>
        <Text style={styles.title}>{tab === 'today' ? 'My rota' : tab === 'timeOff' ? 'Time off' : tab === 'notifications' ? 'Notifications' : 'Profile'}</Text>
      </View>
      <ScrollView contentContainerStyle={styles.content}>
        {tab === 'today' && <>
          <ScrollView horizontal showsHorizontalScrollIndicator={false} contentContainerStyle={styles.dayRow}>
            {dutyDays.sort((a,b)=>(weekdayOrder[a.toLowerCase()] ?? 99)-(weekdayOrder[b.toLowerCase()] ?? 99)).map((day) => (
              <Pressable key={day} onPress={() => setSelectedDay(day)} style={[styles.dayButton, shownDay === day && styles.dayButtonActive]}>
                <Text style={[styles.dayText, shownDay === day && styles.dayTextActive]}>{day.slice(0,3)}</Text>
              </Pressable>
            ))}
          </ScrollView>
          {shownDuties.length === 0 ? <View style={styles.empty}><Text style={styles.emptyTitle}>No sessions assigned</Text><Text style={styles.muted}>Your manager's live rota will appear here.</Text></View> : shownDuties.map((duty) => (
            <View key={duty.id} style={styles.dutyCard}>
              <View style={styles.sessionBadge}><Text style={styles.sessionLabel}>SESSION</Text><Text style={styles.sessionValue}>{sessionNumber(duty.session)}</Text></View>
              <View style={styles.dutyBody}><Text style={styles.activity}>{duty.activity_name}</Text><Text style={styles.groups}>{duty.group_numbers?.length ? `Groups ${duty.group_numbers.join(', ')}` : 'No groups assigned'}</Text></View>
            </View>
          ))}
        </>}

        {tab === 'timeOff' && <>
          <Text style={styles.sectionTitle}>This week</Text>
          {Array.from({ length: 7 }, (_, index) => {
            const date = new Date(); const offset=(date.getDay()+6)%7; date.setDate(date.getDate()-offset+index)
            const key=date.toISOString().slice(0,10); const label=date.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'short'})
            return <View key={key} style={styles.availabilityCard}><Text style={styles.availabilityDate}>{label}</Text><View style={styles.statusRow}>{(['available','holiday','sick'] as Availability[]).map(status => <Pressable key={status} onPress={()=>setDayStatus(key,status)} style={[styles.statusButton,availability[key]===status&&styles.statusActive]}><Text style={[styles.statusText,availability[key]===status&&styles.statusTextActive]}>{status==='available'?'Working':status==='holiday'?'Holiday':'Sick'}</Text></Pressable>)}</View></View>
          })}
        </>}

        {tab === 'notifications' && <View style={styles.empty}><Text style={styles.emptyTitle}>Live updates enabled</Text><Text style={styles.muted}>Rota changes appear automatically while the app is open.</Text></View>}
        {tab === 'profile' && <View style={styles.profileCard}><Text style={styles.profileEmail}>{session.user.email}</Text><Pressable style={styles.signOut} onPress={()=>supabase.auth.signOut()}><Text style={styles.signOutText}>Sign out</Text></Pressable></View>}
      </ScrollView>
      <View style={styles.tabs}>
        {([['today','Today'],['timeOff','Time Off'],['notifications','Notifications'],['profile','Profile']] as [Tab,string][]).map(([key,label])=><Pressable key={key} onPress={()=>setTab(key)} style={styles.tab}><Text style={[styles.tabText,tab===key&&styles.tabActive]}>{label}</Text></Pressable>)}
      </View>
    </SafeAreaView>
  )
}

const styles = StyleSheet.create({
  shell:{flex:1,backgroundColor:'#f4f7f5'}, center:{flex:1,alignItems:'center',justifyContent:'center'}, login:{flex:1,justifyContent:'center',padding:28,backgroundColor:'#0c3b35'}, kicker:{fontSize:12,fontWeight:'800',letterSpacing:1.5,color:'#84d7b0'}, loginTitle:{fontSize:34,fontWeight:'900',color:'white',marginTop:8}, loginCopy:{fontSize:16,lineHeight:23,color:'#c7ddd6',marginVertical:18}, input:{backgroundColor:'#123f3a',borderWidth:1,borderColor:'#35655c',borderRadius:14,padding:16,color:'white',fontSize:16,marginBottom:12}, primary:{backgroundColor:'#53d68d',borderRadius:14,padding:16,alignItems:'center'}, primaryText:{fontWeight:'900',fontSize:16,color:'#07372f'}, header:{backgroundColor:'#0c3b35',paddingHorizontal:20,paddingTop:18,paddingBottom:20}, title:{fontSize:28,fontWeight:'900',color:'white',marginTop:4}, content:{padding:16,paddingBottom:32}, dayRow:{gap:8,paddingBottom:16}, dayButton:{width:54,height:48,borderRadius:14,backgroundColor:'white',alignItems:'center',justifyContent:'center',borderWidth:1,borderColor:'#dce5e0'}, dayButtonActive:{backgroundColor:'#0c3b35'}, dayText:{fontWeight:'800',color:'#39554c'}, dayTextActive:{color:'white'}, dutyCard:{backgroundColor:'white',borderRadius:18,padding:15,marginBottom:12,flexDirection:'row',borderWidth:1,borderColor:'#e1e8e4'}, sessionBadge:{width:70,backgroundColor:'#e7f6ee',borderRadius:13,alignItems:'center',justifyContent:'center',padding:9}, sessionLabel:{fontSize:9,fontWeight:'900',color:'#37705b'}, sessionValue:{fontSize:25,fontWeight:'900',color:'#0c3b35'}, dutyBody:{flex:1,justifyContent:'center',paddingLeft:14}, activity:{fontSize:19,fontWeight:'900',color:'#18352c'}, groups:{fontSize:14,color:'#63776f',marginTop:4}, empty:{backgroundColor:'white',borderRadius:18,padding:24,alignItems:'center'}, emptyTitle:{fontSize:19,fontWeight:'900',color:'#18352c',marginBottom:6}, muted:{color:'#6c7d77',textAlign:'center',lineHeight:20}, sectionTitle:{fontSize:20,fontWeight:'900',color:'#18352c',marginBottom:12}, availabilityCard:{backgroundColor:'white',borderRadius:16,padding:14,marginBottom:10,borderWidth:1,borderColor:'#e1e8e4'}, availabilityDate:{fontWeight:'900',color:'#18352c',marginBottom:10}, statusRow:{flexDirection:'row',gap:7}, statusButton:{flex:1,paddingVertical:10,borderRadius:10,alignItems:'center',backgroundColor:'#eef2f0'}, statusActive:{backgroundColor:'#0c3b35'}, statusText:{fontWeight:'800',fontSize:12,color:'#536961'}, statusTextActive:{color:'white'}, profileCard:{backgroundColor:'white',borderRadius:18,padding:20}, profileEmail:{fontSize:17,fontWeight:'800',color:'#18352c'}, signOut:{marginTop:18,borderWidth:1,borderColor:'#b42318',borderRadius:12,padding:13,alignItems:'center'}, signOutText:{fontWeight:'900',color:'#b42318'}, tabs:{flexDirection:'row',backgroundColor:'white',borderTopWidth:1,borderTopColor:'#dde5e1',paddingVertical:10}, tab:{flex:1,alignItems:'center',paddingVertical:8}, tabText:{fontSize:11,fontWeight:'800',color:'#71827b'}, tabActive:{color:'#0c3b35'},
})
