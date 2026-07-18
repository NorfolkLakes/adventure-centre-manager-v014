import { useEffect, useMemo, useState } from 'react'
import { Building2, CalendarDays, Maximize2, RefreshCw } from 'lucide-react'
import { supabase } from './lib/supabase'
import type { ArrivalAssignment, ProgrammeImport, ProgrammeRow, StaffMember } from './types'

type LiveState = { programme: ProgrammeImport | null; staff: StaffMember[]; arrivalAssignments: Record<string, ArrivalAssignment> }
const buildings = ['Kingfisher','Swan','Grebe','Bittern','Mallard','Teal']
function isoDate(date: Date){return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`}
function weekdayRank(value:string){const k=value.toUpperCase().replace(/[^A-Z]/g,'');return ({MON:1,MONDAY:1,TUE:2,TUES:2,TUESDAY:2,WED:3,WEDNESDAY:3,THU:4,THUR:4,THURS:4,THURSDAY:4,FRI:5,FRIDAY:5,SAT:6,SATURDAY:6,SUN:7,SUNDAY:7} as Record<string,number>)[k]??99}
function dateForDay(programme:ProgrammeImport,day:string){if(/^\d{4}-\d{2}-\d{2}$/.test(day))return day;if(!programme.startDate)return '';const start=new Date(`${programme.startDate}T12:00:00`);const rank=weekdayRank(day);if(Number.isNaN(start.getTime())||rank===99)return '';const startRank=((start.getDay()+6)%7)+1;start.setDate(start.getDate()+((rank-startRank+7)%7));return isoDate(start)}
function arrivalKey(row:ProgrammeRow){return `${row.day}::${row.id}`}
function accommodation(flatIds:string[]=[]){const names=[...new Set(flatIds.map(id=>buildings[Number(id.split('-')[0])-1]).filter(Boolean))];return names.join(', ')||'Not allocated'}

export default function ReceptionDisplay({onExit}:{onExit:()=>void}){
 const [state,setState]=useState<LiveState>({programme:null,staff:[],arrivalAssignments:{}});const [now,setNow]=useState(new Date())
 async function load(){const {data}=await supabase.from('app_live_state').select('state').eq('id','main').maybeSingle();if(data?.state)setState(data.state as LiveState)}
 useEffect(()=>{load();const timer=window.setInterval(()=>{setNow(new Date());load()},30000);const channel=supabase.channel('reception-display-live').on('postgres_changes',{event:'*',schema:'public',table:'app_live_state',filter:'id=eq.main'},load).subscribe();return()=>{clearInterval(timer);supabase.removeChannel(channel)}},[])
 const today=isoDate(now)
 const arrivalRows=useMemo(()=>state.programme?.rows.filter(r=>r.session==='3'&&dateForDay(state.programme!,r.day)===today&&r.schoolLabel?.trim())??[],[state.programme,today])
 const cards=useMemo(()=>arrivalRows.map(row=>{const details=state.arrivalAssignments[arrivalKey(row)]??{guideIds:[]};const groups=[...new Set(row.cells.map(c=>c.group))].sort((a,b)=>a-b);return {name:row.schoolLabel||'School',groups,building:accommodation(details.flatIds),leader:state.staff.find(s=>s.id===details.leaderId)?.name||'Not assigned'}}),[arrivalRows,state.arrivalAssignments,state.staff])
 const onSite=useMemo(()=>state.programme?.schoolDetails?.filter(s=>s.arrivalDate<=today&&s.departureDate>=today)??[],[state.programme,today])
 async function fullScreen(){try{await document.documentElement.requestFullscreen()}catch{}}
 return <main className="staff-display-shell reception-display-shell"><header className="staff-display-header"><div className="staff-display-brand"><img src={`${import.meta.env.BASE_URL}manor-adventure-logo.png`} alt="Manor Adventure"/><div><p>NORFOLK LAKES</p><h1>Reception Display</h1></div></div><div className="staff-display-date"><CalendarDays/><strong>{now.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long',year:'numeric'})}</strong><span>{now.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'})}</span></div><div className="staff-display-actions"><button onClick={()=>void load()}><RefreshCw/>Refresh</button><button onClick={fullScreen}><Maximize2/>Full screen</button><button onClick={onExit}>Exit display</button></div></header>
 <section className="reception-display-content"><div className="reception-block"><h2>Schools arriving today</h2>{cards.length?<div className="reception-card-grid">{cards.map(card=><article className="reception-school-card" key={card.name}><h3>{card.name}</h3><dl><div><dt>Groups</dt><dd>{card.groups.length?card.groups.map(g=>`G${g}`).join(', '):'Not detected'}</dd></div><div><dt>Building</dt><dd>{card.building}</dd></div><div><dt>School Leader</dt><dd>{card.leader}</dd></div></dl></article>)}</div>:<div className="staff-display-empty compact"><h3>No schools arriving today</h3></div>}</div>
 <div className="reception-block"><h2>Schools on site</h2><div className="building-board">{onSite.length?onSite.map(s=>{const row=state.programme?.rows.find(r=>r.schoolLabel===s.schoolName);const details=row?state.arrivalAssignments[arrivalKey(row)]:undefined;return <article key={s.id}><Building2/><div><strong>{s.schoolName}</strong><span>{accommodation(details?.flatIds)}</span></div></article>}):<p>No schools are currently on site.</p>}</div></div></section><footer className="staff-display-footer">Live read-only reception display · refreshes automatically every 30 seconds</footer></main>
}
