export type ProgrammePlacement = { key: string; code: string }

export type ProgrammeTask = {
  id: string
  candidates: ProgrammePlacement[][]
}

export type ProgrammeTaskSolverOptions = {
  initialAssignments: Record<string, string>
  tasks: ProgrammeTask[]
  capacityForCode: (code: string) => number
  slotForKey?: (key: string) => string
}

export type ProgrammeTaskSolverResult = {
  success: boolean
  assignments: Record<string, string>
  unresolvedTaskIds: string[]
}

function defaultSlotForKey(key: string) {
  const [day, session] = key.split('|')
  return `${day}|${session}`
}

export function solveProgrammeTasks(options: ProgrammeTaskSolverOptions): ProgrammeTaskSolverResult {
  const assignments = { ...options.initialAssignments }
  const slotForKey = options.slotForKey ?? defaultSlotForKey
  const remaining = [...options.tasks]

  const canApply = (candidate: ProgrammePlacement[]) => {
    const keys = new Set<string>()
    const additions = new Map<string, number>()

    for (const placement of candidate) {
      if (keys.has(placement.key)) return false
      keys.add(placement.key)
      const current = assignments[placement.key] ?? ''
      if (current && current !== placement.code) return false
      if (current === placement.code) continue
      const capacityKey = `${slotForKey(placement.key)}|${placement.code}`
      additions.set(capacityKey, (additions.get(capacityKey) ?? 0) + 1)
    }

    for (const [capacityKey, addition] of additions) {
      const parts = capacityKey.split('|')
      const code = parts.pop() ?? ''
      const slot = parts.join('|')
      let running = 0
      for (const [key, value] of Object.entries(assignments)) {
        if (value === code && slotForKey(key) === slot) running += 1
      }
      if (running + addition > Math.max(1, options.capacityForCode(code))) return false
    }
    return true
  }

  const apply = (candidate: ProgrammePlacement[]) => {
    const changed: string[] = []
    for (const placement of candidate) {
      if ((assignments[placement.key] ?? '') === placement.code) continue
      assignments[placement.key] = placement.code
      changed.push(placement.key)
    }
    return changed
  }

  const undo = (changed: string[]) => {
    for (const key of changed) assignments[key] = ''
  }

  const search = (): boolean => {
    if (!remaining.length) return true

    let selectedIndex = -1
    let selectedCandidates: ProgrammePlacement[][] = []
    for (let index = 0; index < remaining.length; index += 1) {
      const feasible = remaining[index].candidates.filter(canApply)
      if (!feasible.length) return false
      if (selectedIndex < 0 || feasible.length < selectedCandidates.length) {
        selectedIndex = index
        selectedCandidates = feasible
        if (feasible.length === 1) break
      }
    }

    const [task] = remaining.splice(selectedIndex, 1)
    for (const candidate of selectedCandidates) {
      const changed = apply(candidate)
      if (search()) return true
      undo(changed)
    }
    remaining.splice(selectedIndex, 0, task)
    return false
  }

  const success = search()
  return {
    success,
    assignments: success ? assignments : { ...options.initialAssignments },
    unresolvedTaskIds: success ? [] : remaining.map((task) => task.id),
  }
}

export type RotationGroupPlan = {
  groupId: string
  slotKeys: string[]
  activityCodes: string[]
}

export function solveRotatingGroupActivities(options: {
  initialAssignments: Record<string, string>
  groups: RotationGroupPlan[]
  capacityForCode: (code: string) => number
  slotForKey?: (key: string) => string
}): ProgrammeTaskSolverResult {
  const assignments = { ...options.initialAssignments }
  const slotForKey = options.slotForKey ?? defaultSlotForKey

  const running = (key: string, code: string) => {
    const targetSlot = slotForKey(key)
    let count = 0
    for (const [assignmentKey, value] of Object.entries(assignments)) {
      if (value === code && slotForKey(assignmentKey) === targetSlot) count += 1
    }
    return count
  }

  const arrangementsFor = (plan: RotationGroupPlan) => {
    const n = plan.slotKeys.length
    if (n !== plan.activityCodes.length) return []
    if (!n) return [[]]
    const bases = [plan.activityCodes, [...plan.activityCodes].reverse()]
    const seen = new Set<string>()
    const results: ProgrammePlacement[][] = []
    for (const base of bases) {
      for (let offset = 0; offset < n; offset += 1) {
        const candidate = plan.slotKeys.map((key, index) => ({ key, code: base[(index + offset) % n] }))
        const signature = candidate.map((item) => item.code).join('|')
        if (seen.has(signature)) continue
        seen.add(signature)
        results.push(candidate)
      }
    }
    return results
  }

  const groups = options.groups.map((plan) => ({ plan, candidates: arrangementsFor(plan) }))
  if (groups.some((entry) => !entry.candidates.length)) {
    return { success: false, assignments: { ...options.initialAssignments }, unresolvedTaskIds: groups.filter((entry) => !entry.candidates.length).map((entry) => entry.plan.groupId) }
  }

  const search = (index: number): boolean => {
    if (index >= groups.length) return true
    const entry = groups[index]
    const candidates = [...entry.candidates].sort((a, b) => {
      const pressure = (candidate: ProgrammePlacement[]) => candidate.reduce((total, placement) => total + running(placement.key, placement.code) / Math.max(1, options.capacityForCode(placement.code)), 0)
      return pressure(a) - pressure(b)
    })

    for (const candidate of candidates) {
      let valid = true
      for (const placement of candidate) {
        const current = assignments[placement.key] ?? ''
        if ((current && current !== placement.code) || (!current && running(placement.key, placement.code) >= Math.max(1, options.capacityForCode(placement.code)))) {
          valid = false
          break
        }
      }
      if (!valid) continue
      const changed: string[] = []
      for (const placement of candidate) {
        if ((assignments[placement.key] ?? '') === placement.code) continue
        assignments[placement.key] = placement.code
        changed.push(placement.key)
      }
      if (search(index + 1)) return true
      for (const key of changed) assignments[key] = ''
    }
    return false
  }

  const success = search(0)
  return {
    success,
    assignments: success ? assignments : { ...options.initialAssignments },
    unresolvedTaskIds: success ? [] : groups.map((entry) => entry.plan.groupId),
  }
}
