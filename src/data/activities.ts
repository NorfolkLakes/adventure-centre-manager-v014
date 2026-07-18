import type { Activity } from '../types'

export const activityNames: Record<string, string> = {
  "SAIL LEAD": "Sailing Lead",
  "SAIL PB": "Sailing Powerboat",
  "SAIL A": "Sailing Assistant",
  "CANOE LEAD": "Canoe Lead",
  "KAYAK LEAD": "Kayak Lead",
  "CANOE": "Canoeing",
  "KAYAK": "Kayaking",
  "RAFT": "Raft Building",
  "SUP": "Stand-Up Paddleboarding",
  "GSUP": "Group SUP",
  "GCAN": "Group Canoe",
  "HR UP": "High Ropes \u2013 Up",
  "HR GR": "High Ropes \u2013 Ground",
  "CLIMB": "Climbing",
  "CS": "Crate Stack",
  "RIFLES": "Rifles",
  "ARCH": "Archery",
  "FENCE": "Fencing",
  "AERO": "Aerial Activity",
  "AXE": "Axe Throwing",
  "BOULD": "Bouldering",
  "BT": "Blind Trail",
  "CAVE": "Caving",
  "LR": "Low Ropes",
  "OC": "Obstacle Course",
  "SURV": "Survival",
  "BIVI": "Bivouac",
  "IES": "Initiative Exercises",
  "LAKE WALK": "Lake Walk",
  "QUIZ": "Quiz",
  "TOW": "Tower",
  "VIDEO": "Video",
  "VB": "Volleyball",
  "ORIENT": "Orienteering",
  "SCAV": "Scavenger Hunt",
  "MO": "Manor Olympics",
  "TG": "Team Games",
  "WG": "Wide Games",
  "CF": "Campfire",
  "DISCO": "Disco",
  "LOW WALK": "Low Walk",
  "FILM": "Film",
  "SAIL": "Sailing",
  "HR": "High Ropes",
  "Z": "Unavailable / No activity"
}

export const startingActivities: Activity[] = Object.entries(activityNames).map(
  ([code, name]) => ({ code, name, colour: '#dce8f5', equipmentQuantity: 0, capacity: 1, enabled: true, notes: '' }),
)

export function activityNameFromList(
  activities: Activity[],
  code: string,
) {
  return (
    activities.find((activity) => activity.code === code.toUpperCase())?.name ??
    code
  )
}
