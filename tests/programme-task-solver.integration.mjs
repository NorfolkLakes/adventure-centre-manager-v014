import assert from 'node:assert/strict'
import { solveProgrammeTasks } from '../src/programmeTaskSolver.ts'

const days = ['Mon','Tue','Wed','Thu','Fri']
const sessions = ['1','2','3','4','5']
const slot = (day, session, group) => `${day}|${session}|${group}`
const task = (id, code, candidates) => ({ id, candidates: candidates.map((key) => [{ key, code }]) })
const blockTask = (id, code, blocks, group) => ({ id, candidates: blocks.map(([day,a,b]) => [{ key: slot(day,a,group), code }, { key: slot(day,b,group), code }]) })
const pairTask = (id, blocks, group) => ({ id, candidates: blocks.flatMap(([day,a,b]) => [[{key:slot(day,a,group),code:'CANOE'},{key:slot(day,b,group),code:'KAYAK'}],[{key:slot(day,a,group),code:'KAYAK'},{key:slot(day,b,group),code:'CANOE'}]]) })

const capacities = { CANOE: 3, KAYAK: 3, RAFT: 2, CLIMB: 2, ARCH: 2, RIFLES: 2, CF: 30, DISCO: 30 }
const capacityForCode = (code) => capacities[code] ?? 30

function schoolSlots(arrival, departure, group) {
  const ai = days.indexOf(arrival), di = days.indexOf(departure)
  const keys = []
  for (let i=ai;i<=di;i++) {
    const day = days[i]
    for (const s of sessions) {
      if (i===ai && ['1','2','3'].includes(s)) continue
      if (i===di && ['3','4','5'].includes(s)) continue
      keys.push(slot(day,s,group))
    }
  }
  return keys
}

function addSchoolTasks(tasks, {id, groups, arrival, departure, activities, evening, waterMode}) {
  const allSlots = Object.fromEntries(groups.map(g => [g, schoolSlots(arrival, departure, g)]))
  const blocks = days.slice(days.indexOf(arrival), days.indexOf(departure)+1).flatMap(day => [['1','2'],['3','4']].map(([a,b]) => [day,a,b])).filter(([day,a,b]) => allSlots[groups[0]].includes(slot(day,a,groups[0])) && allSlots[groups[0]].includes(slot(day,b,groups[0])))
  if (evening) {
    const candidates = days.slice(days.indexOf(arrival), days.indexOf(departure)+1).filter(day => groups.every(g => allSlots[g].includes(slot(day,'5',g)))).map(day => groups.map(g => ({key:slot(day,'5',g),code:evening})))
    tasks.push({id:`${id}-evening`, candidates})
  }
  for (const g of groups) {
    if (waterMode === 'super') {
      tasks.push(blockTask(`${id}-g${g}-canoe`,'CANOE',blocks,g))
      tasks.push(blockTask(`${id}-g${g}-kayak`,'KAYAK',blocks,g))
    } else if (waterMode === 'bargain') {
      tasks.push(pairTask(`${id}-g${g}-pair`,blocks,g))
    }
    const reservedCodes = new Set(waterMode ? ['CANOE','KAYAK'] : [])
    const other = activities.filter(a => !reservedCodes.has(a) && a !== evening)
    for (const code of other) tasks.push(task(`${id}-g${g}-${code}`, code, allSlots[g]))
  }
}

const tasks = []
// Monday-Friday, 5 days, four groups, Outdoor Challenge style water blocks.
addSchoolTasks(tasks, {id:'five-day', groups:[1,2,3,4], arrival:'Mon', departure:'Fri', activities:['CANOE','KAYAK','RAFT','CLIMB','ARCH','RIFLES','FENCING','OC','LR','BT','MO','AIR','QUIZ','TG','TOW','CAVE','SUP'], evening:'CF', waterMode:'super'})
// Monday-Wednesday, 3 days, three groups, Bargain water pair.
addSchoolTasks(tasks, {id:'mon-three', groups:[5,6,7], arrival:'Mon', departure:'Wed', activities:['CANOE','KAYAK','RAFT','ARCH','RIFLES','FENCING','OC','LR'], evening:null, waterMode:'bargain'})
// Wednesday-Friday, 3 days, four groups, different activities and Disco.
addSchoolTasks(tasks, {id:'wed-three', groups:[8,9,10,11], arrival:'Wed', departure:'Fri', activities:['RAFT','CLIMB','ARCH','RIFLES','FENCING','OC','LR','DISCO'], evening:'DISCO', waterMode:null})

const result = solveProgrammeTasks({ initialAssignments:{}, tasks, capacityForCode })
assert.equal(result.success, true, `solver failed: ${result.unresolvedTaskIds.join(', ')}`)

for (const t of tasks) {
  assert.ok(t.candidates.some(candidate => candidate.every(p => result.assignments[p.key] === p.code)), `task not satisfied: ${t.id}`)
}

const bySlotCode = new Map()
for (const [key, code] of Object.entries(result.assignments)) {
  const [day, session] = key.split('|')
  const k = `${day}|${session}|${code}`
  bySlotCode.set(k, (bySlotCode.get(k) ?? 0) + 1)
}
for (const [k,count] of bySlotCode) {
  const code = k.split('|').at(-1)
  assert.ok(count <= capacityForCode(code), `${k} exceeded capacity ${count}/${capacityForCode(code)}`)
}

// Run it a second time to catch non-deterministic or mutation bugs.
const second = solveProgrammeTasks({ initialAssignments:{}, tasks, capacityForCode })
assert.equal(second.success, true)
assert.deepEqual(second.assignments, result.assignments)
console.log(`PASS: ${tasks.length} real scheduling tasks across 3 overlapping schools; repeated run identical.`)
