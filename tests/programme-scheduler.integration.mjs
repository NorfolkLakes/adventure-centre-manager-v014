import assert from 'node:assert/strict'
import { solveProgrammeTasks, solveRotatingGroupActivities } from '../src/programmeTaskSolver.ts'

const capacities = { CANOE: 3, KAYAK: 3, RAFT: 2, CLIMB: 2, ARCH: 2, RIFLES: 2, CF: 30, DISCO: 30 }
const cap = (code) => capacities[code] ?? 30
const days = ['Mon','Tue','Wed','Thu','Fri']
const sessions = ['1','2','3','4','5']
const key = (day,session,group) => `${day}|${session}|${group}`

function slots(arrival, departure, group) {
  const a=days.indexOf(arrival), d=days.indexOf(departure), out=[]
  for(let i=a;i<=d;i++) for(const s of sessions){
    if(i===a && ['1','2','3'].includes(s)) continue
    if(i===d && ['3','4','5'].includes(s)) continue
    out.push(key(days[i],s,group))
  }
  return out
}

function blocks(arrival, departure, group) {
  const available = new Set(slots(arrival, departure, group))
  return days.slice(days.indexOf(arrival),days.indexOf(departure)+1).flatMap(day => [['1','2'],['3','4']].map(([a,b]) => [key(day,a,group),key(day,b,group)])).filter(([a,b]) => available.has(a)&&available.has(b))
}

function scheduleSchool(existing, spec) {
  const special=[]
  const reserved = new Map(spec.groups.map(g => [g,new Set()]))
  if(spec.evening){
    const cands = days.slice(days.indexOf(spec.arrival),days.indexOf(spec.departure)+1)
      .filter(day => spec.groups.every(g => slots(spec.arrival,spec.departure,g).includes(key(day,'5',g))))
      .map(day => spec.groups.map(g => ({key:key(day,'5',g),code:spec.evening})))
    special.push({id:`${spec.name}-evening`,candidates:cands})
    for(const g of spec.groups) reserved.get(g).add(spec.evening)
  }
  if(spec.water==='super'){
    for(const g of spec.groups){
      for(const code of ['CANOE','KAYAK']){
        special.push({id:`${spec.name}-g${g}-${code}`,candidates:blocks(spec.arrival,spec.departure,g).map(([a,b])=>[{key:a,code},{key:b,code}])})
        reserved.get(g).add(code)
      }
    }
  } else if(spec.water==='bargain'){
    for(const g of spec.groups){
      special.push({id:`${spec.name}-g${g}-pair`,candidates:blocks(spec.arrival,spec.departure,g).flatMap(([a,b])=>[[{key:a,code:'CANOE'},{key:b,code:'KAYAK'}],[{key:a,code:'KAYAK'},{key:b,code:'CANOE'}]])})
      reserved.get(g).add('CANOE'); reserved.get(g).add('KAYAK')
    }
  }
  const s1=solveProgrammeTasks({initialAssignments:existing,tasks:special,capacityForCode:cap})
  assert.equal(s1.success,true,`${spec.name} special tasks failed`)
  const groups=spec.groups.map(g=>{
    const slotKeys=slots(spec.arrival,spec.departure,g).filter(k=>!s1.assignments[k])
    const activityCodes=spec.activities.filter(code=>!reserved.get(g).has(code))
    assert.equal(slotKeys.length,activityCodes.length,`${spec.name} G${g} count mismatch ${slotKeys.length}/${activityCodes.length}`)
    return {groupId:`${spec.name}-G${g}`,slotKeys,activityCodes}
  })
  const s2=solveRotatingGroupActivities({initialAssignments:s1.assignments,groups,capacityForCode:cap})
  assert.equal(s2.success,true,`${spec.name} rotation failed`)
  return s2.assignments
}

function validate(assignments,spec){
  for(const g of spec.groups){
    const values=slots(spec.arrival,spec.departure,g).map(k=>assignments[k])
    assert.ok(values.every(Boolean),`${spec.name} G${g} has blank`)
    const counts=new Map(); for(const v of values) counts.set(v,(counts.get(v)||0)+1)
    for(const code of spec.activities){
      const expected = spec.water==='super' && ['CANOE','KAYAK'].includes(code) ? 2 : 1
      assert.equal(counts.get(code),expected,`${spec.name} G${g} ${code} count`)
    }
  }
}

const fiveDay={name:'five-day',arrival:'Mon',departure:'Fri',groups:[1,2,3,4],water:'super',evening:'CF',activities:['CANOE','CANOE','KAYAK','KAYAK','CF','RAFT','CLIMB','ARCH','RIFLES','FENCING','OC','LR','BT','MO','AIR','QUIZ','TG','TOW','CAVE']}
const mon3={name:'mon-three',arrival:'Mon',departure:'Wed',groups:[5,6,7],water:'bargain',activities:['CANOE','KAYAK','RAFT','ARCH','RIFLES','FENCING','OC','LR','BT']}
const wed3={name:'wed-three',arrival:'Wed',departure:'Fri',groups:[8,9,10,11],evening:'DISCO',activities:['DISCO','RAFT','CLIMB','ARCH','RIFLES','FENCING','OC','LR','BT']}

let assignments={}
for(const spec of [fiveDay,mon3,wed3]){ assignments=scheduleSchool(assignments,spec); validate(assignments,spec) }
for(const [slotCode,count] of Object.entries(Object.entries(assignments).reduce((m,[k,v])=>{const [d,s]=k.split('|');const id=`${d}|${s}|${v}`;m[id]=(m[id]||0)+1;return m},{}))){
  const code=slotCode.split('|').at(-1); assert.ok(count<=cap(code),`${slotCode} capacity ${count}/${cap(code)}`)
}

let assignments2={}
for(const spec of [fiveDay,mon3,wed3]) assignments2=scheduleSchool(assignments2,spec)
assert.deepEqual(assignments2,assignments)
console.log('PASS: Monday 5-day + Monday 3-day + Wednesday 3-day, different activities, capacities, water blocks, evening activities, repeated twice.')
