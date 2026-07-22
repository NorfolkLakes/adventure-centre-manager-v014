#!/usr/bin/env python3
from pathlib import Path
import re
import sys

path = Path("src/ManagerApp.tsx")
if not path.exists():
    path = Path("ManagerApp.tsx")

if not path.exists():
    raise SystemExit("Could not find src/ManagerApp.tsx or ManagerApp.tsx. Run this from the project folder.")

text = path.read_text(encoding="utf-8")
original = text

old_block = """  const schoolsOnSite = new Set(programme?.rows.map(arrivalSchoolName).filter(Boolean) ?? []).size
  const availableTodayCount = activeStaffingDay
    ? (workingByDay[activeStaffingDay] ?? staff.map((m) => m.id)).filter((id) => !unavailableStaffIdsForDay(activeStaffingDay).has(id)).length
    : staff.length
"""

new_block = """  const schoolsOnSite = new Set(programme?.rows.map(arrivalSchoolName).filter(Boolean) ?? []).size

  // Keep the dashboard's live staff count current as the working day changes.
  const [dashboardNow, setDashboardNow] = useState(() => new Date())
  useEffect(() => {
    const timer = window.setInterval(() => setDashboardNow(new Date()), 60_000)
    return () => window.clearInterval(timer)
  }, [])

  const availableTodayCount = useMemo(() => {
    const londonParts = new Intl.DateTimeFormat('en-GB', {
      timeZone: 'Europe/London',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      weekday: 'short',
      hour: '2-digit',
      minute: '2-digit',
      hourCycle: 'h23',
    }).formatToParts(dashboardNow)

    const part = (type: Intl.DateTimeFormatPartTypes) =>
      londonParts.find((item) => item.type === type)?.value ?? ''

    const hour = Number(part('hour'))
    const minute = Number(part('minute'))
    const minutesNow = hour * 60 + minute

    // The operational working day is 08:30–20:30.
    if (minutesNow < 8 * 60 + 30 || minutesNow > 20 * 60 + 30) return 0

    const todayIso = `${part('year')}-${part('month')}-${part('day')}`
    const weekdayShort = part('weekday').toUpperCase()
    const weekdayAliases: Record<string, string[]> = {
      MON: ['MON', 'MONDAY'],
      TUE: ['TUE', 'TUES', 'TUESDAY'],
      WED: ['WED', 'WEDNESDAY'],
      THU: ['THU', 'THUR', 'THURS', 'THURSDAY'],
      FRI: ['FRI', 'FRIDAY'],
      SAT: ['SAT', 'SATURDAY'],
      SUN: ['SUN', 'SUNDAY'],
    }
    const aliases = weekdayAliases[weekdayShort] ?? [weekdayShort]
    const normaliseDay = (value: string) => value.trim().toUpperCase().replace(/[^A-Z0-9-]/g, '')
    const isTodayValue = (value: string) => {
      const normalised = normaliseDay(value)
      return value === todayIso || aliases.some((alias) => normalised === alias)
    }

    const programmeDay = programmeDays.find((day) => isTodayValue(day))
    const scheduledIds = programmeDay && workingByDay[programmeDay]
      ? new Set(workingByDay[programmeDay])
      : null

    return staff.filter((member) => {
      // When the day has an explicit working list, only count staff on that list.
      if (scheduledIds && !scheduledIds.has(member.id)) return false

      const entry = daysOff.find(
        (item) => item.staff_id === member.id && isTodayValue(item.day),
      )
      const status: DayOffStatus | 'working' = entry?.status ?? 'working'

      if (status === 'off' || status === 'hol' || status === 'sick') return false
      if (status === 'am_off') return minutesNow >= 12 * 60 + 15
      if (status === 'pm_off') return minutesNow <= 17 * 60 + 15
      return true
    }).length
  }, [dashboardNow, staff, daysOff, programmeDays, workingByDay])
"""

if old_block not in text:
    raise SystemExit(
        "The expected dashboard count block was not found. "
        "Make sure this is the latest ManagerApp.tsx used by your project."
    )

text = text.replace(old_block, new_block, 1)
text = text.replace("<small>staff available today</small>", "<small>staff currently on site</small>", 1)
text = text.replace('label="Staff available"', 'label="Staff currently on site"', 1)

# Update the visible version badge without disturbing other version strings.
text = re.sub(
    r'(<span className="release-pill">)v[^<]+(</span>)',
    r'\1v3.1\2',
    text,
    count=1,
)

backup = path.with_suffix(path.suffix + ".before-v3.1")
backup.write_text(original, encoding="utf-8")
path.write_text(text, encoding="utf-8")

print(f"Updated: {path}")
print(f"Backup:  {backup}")
print("")
print("Next run:")
print("  npm run build")
