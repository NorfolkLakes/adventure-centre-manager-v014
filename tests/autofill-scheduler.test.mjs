import assert from 'node:assert/strict'

function weightedCount(codes, purchaseType) {
  return codes.reduce((total, code) => total + ((purchaseType === 'super' || purchaseType === 'outdoor') && (code === 'CANOE' || code === 'KAYAK') ? 2 : 1), 0)
}

function buildCandidates({ requested, backups, protectedCodes, slotCount, purchaseType }) {
  const protectedSet = new Set(protectedCodes)
  const protectedRequested = requested.filter(code => protectedSet.has(code))
  const optional = requested.filter(code => !protectedSet.has(code))
  const usableBackups = backups.filter(Boolean).filter(code => !requested.includes(code))
  const all = [...optional, ...usableBackups]
  const candidates = []
  for (let mask = 0; mask < (1 << all.length); mask++) {
    const codes = [...protectedRequested]
    for (let i = 0; i < all.length; i++) if (mask & (1 << i)) codes.push(all[i])
    if (new Set(codes).size !== codes.length) continue
    if (weightedCount(codes, purchaseType) === slotCount) candidates.push(codes)
  }
  return candidates
}

function solve({ groups, slots, codes, capacityByCode, allowed }) {
  const assignments = new Map()
  const remaining = new Map()
  for (const group of groups) for (const code of codes) remaining.set(`${group}|${code}`, { group, code, count: 1 })

  const running = (slot, code) => groups.filter(group => assignments.get(`${group}|${slot}`) === code).length
  const total = () => [...remaining.values()].reduce((n, item) => n + item.count, 0)

  function recurse() {
    if (!total()) return true
    let selectedKey = ''
    let selectedSlots = []
    for (const [key, item] of remaining) {
      if (!item.count) continue
      const candidates = slots.filter(slot => !assignments.has(`${item.group}|${slot}`)
        && allowed(item.group, slot, item.code)
        && running(slot, item.code) < (capacityByCode[item.code] ?? groups.length))
      if (!candidates.length) return false
      if (!selectedKey || candidates.length < selectedSlots.length) {
        selectedKey = key
        selectedSlots = candidates
      }
    }
    const item = remaining.get(selectedKey)
    for (const slot of selectedSlots) {
      assignments.set(`${item.group}|${slot}`, item.code)
      item.count--
      if (recurse()) return true
      item.count++
      assignments.delete(`${item.group}|${slot}`)
    }
    return false
  }

  return { ok: recurse(), assignments }
}

// Case 1: nine available sessions and nine one-session activities.
{
  const groups = [1, 2, 3, 4]
  const slots = Array.from({ length: 9 }, (_, i) => `S${i + 1}`)
  const codes = Array.from({ length: 9 }, (_, i) => `A${i + 1}`)
  const result = solve({ groups, slots, codes, capacityByCode: {}, allowed: () => true })
  assert.equal(result.ok, true)
  for (const group of groups) {
    const values = slots.map(slot => result.assignments.get(`${group}|${slot}`))
    assert.equal(values.filter(Boolean).length, 9)
    assert.equal(new Set(values).size, 9)
  }
}

// Case 2: eight requested activities plus one selected backup makes nine sessions.
{
  const requested = ['A1','A2','A3','A4','A5','A6','A7','A8']
  const candidates = buildCandidates({ requested, backups: ['BACKUP'], protectedCodes: [], slotCount: 9, purchaseType: 'bargain' })
  assert.ok(candidates.some(codes => codes.includes('BACKUP') && codes.length === 9))
}

// Case 3: package weighting can exceed the slot count; optional activities are trimmed,
// while protected Canoe/Kayak remain in the plan.
{
  const requested = ['CANOE','KAYAK','A1','A2','A3','A4','A5','A6','A7']
  const candidates = buildCandidates({ requested, backups: [], protectedCodes: ['CANOE','KAYAK'], slotCount: 9, purchaseType: 'super' })
  assert.ok(candidates.length > 0)
  assert.ok(candidates.every(codes => codes.includes('CANOE') && codes.includes('KAYAK')))
  assert.ok(candidates.every(codes => weightedCount(codes, 'super') === 9))
}

// Case 4: constrained layout that needs whole-programme backtracking.
{
  const groups = [1, 2]
  const slots = ['S1','S2','S3']
  const codes = ['A','B','C']
  const allowed = (group, slot, code) => {
    if (group === 1 && code === 'C') return slot === 'S1'
    if (group === 2 && code === 'B') return slot === 'S1'
    return true
  }
  const result = solve({ groups, slots, codes, capacityByCode: { A: 1, B: 1, C: 1 }, allowed })
  assert.equal(result.ok, true)
  for (const group of groups) assert.equal(new Set(slots.map(slot => result.assignments.get(`${group}|${slot}`))).size, 3)
}

console.log('Auto Fill scheduler tests passed: 4/4')
