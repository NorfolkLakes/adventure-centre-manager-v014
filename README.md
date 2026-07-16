# Adventure Centre Manager v0.28

This version adds the first cloud login and staff rota system.

## Included

- Supabase email/password authentication
- Manager and staff roles
- Manager keeps the existing operational app
- Staff receive a phone-friendly My Rota screen
- Staff can see:
  - Day
  - Session
  - Activity
  - Group number
  - School arrival duties
- Staff Management includes a login email for every person
- Manager can publish the current rota to Supabase
- Staff rota updates automatically when the manager republishes
- Row Level Security restricts staff to duties matching their email

## Setup

### 1. Environment

Copy `.env.example` to `.env.local` and insert your Supabase Project URL and publishable key.

### 2. Database

Open Supabase -> SQL Editor -> New query.

Paste and run:

`supabase/setup.sql`

### 3. Create the manager

1. Run the app.
2. Select Create staff account and create your own account.
3. In Supabase -> Table Editor -> profiles, change your role to `manager`.
4. Sign out and sign back in.

### 4. Connect staff accounts

In Staff Management, enter each person's login email.

Each staff member creates an account using that exact email.

### 5. Publish

Complete the staffing rota and click **Publish staff rota**.

## Run

```bash
npm install
npm run dev
```

## Security

The browser uses only the Supabase publishable key. Row Level Security policies in `setup.sql` control access. Never put a secret or service-role key in `.env.local`.


## New in v0.14

- Manager rota changes sync automatically to Supabase
- Publish button replaced with a live cloud-sync status
- Staff My Rota redesigned for phones
- Day tabs make the weekly rota easier to browse
- Duties show clear session, activity, school and group information
- Staff screens refresh automatically when the manager changes the rota
- Excel printing now includes a Staff Room Wall Rota sheet


## v0.20 Commercial operations release

- Live Supabase rota syncing with realtime staff-phone updates
- Staff self-service Available / Holiday / Sick status
- Manager auto-fill consumes staff availability automatically
- Qualification expiry dates and expiry-aware assignment
- Whole-programme AI rota builder with workload balancing and conflict prevention
- Dashboard metrics for schools, availability, shortages and expiring qualifications
- Outstanding equipment-check dashboard

After deploying v0.20, run `supabase/setup.sql` once in the Supabase SQL Editor to add the availability table and policies.


## v0.20.1 Arrival staffing correction
- Party Leader can be a Staff member or Team Leader.
- Manual group assignment prevents cross-school double-booking.
- One instructor can cover at most two groups, both from the same school.
- Later sessions are unrestricted by school arrival assignments.
- Staff login combines two arrival groups into one duty card.

## v0.21 Multi-school arrival staffing

- Visible v0.21 release badge on the manager dashboard.
- Select up to six schools from the programme on Daily Staffing.
- Assign any working staff member or team leader as Party Leader.
- Assign one instructor per group manually or use school-specific auto-fill.
- One instructor can cover no more than two groups, and both groups must be from the same school.
- Staff cannot be assigned to two schools during Session 3.
- Later activity sessions remain unrestricted by arrival-school allocation.
- Staff accounts show school, group(s), and Party Leader or Accommodation / Fire Alarm Instructor role.


## v0.27 Arrival & Accommodation
- One card per arriving school.
- Buildings 1–6, arrival time and departure day/time.
- Party Leader and group assignment with school-specific auto-fill.
- Overlap-aware staff conflict prevention.
- Building occupancy conflict warnings and Accommodation Overview page.
- Staff duties show day, school, building, Party Leader, groups and role “Accommodation”.

## v0.27 Arrivals module
- Detects named schools only from Monday, Wednesday and Friday Session 3 programme rows.
- Removes school-arrival rows from normal Daily Staffing and activity auto-fill.
- Adds a dedicated Arrivals page for building, Party Leader and school-group staffing.
- Keeps normal activity staffing independent after the arrival session.


## v0.27 inline school-name detection
- Detects school names written directly inside Monday, Wednesday or Friday Session 3 group cells.
- Removes those rows from Daily Staffing and sends them to Arrivals.
- Keeps normal Session 3 activity rows in Daily Staffing.


## v0.27 parser separation fix

- Recognises abbreviated programme days (`MON`, `WED`, `FRI`) as well as full day names.
- Classifies non-activity values in Session 3 on arrival days before Daily Staffing is generated.
- School names such as GREAT BRADFORD and HENLOW are sent only to Arrivals.
- Genuine Session 3 activities remain in Daily Staffing.


## v0.27 school group and flat allocation
- Counts each school's programme groups automatically.
- Adds Auto-fill school and Auto-fill all schools after Party Leaders are selected.
- Keeps one instructor with one school during Session 3; a maximum of two groups per instructor.
- Replaces arrival/departure timing with flat allocation.
- Supports Buildings 1–6 with Flats 1–5 in each building.
- A school can use flats across more than one building.
- Prevents the same flat being allocated twice on the same arrival day.


## v0.27 school group range detection
- Reads programme group columns G1 through G30.
- A school name starts a school group block.
- Blank cells after the school name inherit that school until the next school name or recognised activity.
- Arrival cards show the exact group range and total group count automatically.


## v0.28 Random arrivals, staffing alerts and school notes

- Arrival auto-fill shuffles eligible staff on every run rather than assigning alphabetically or by role order.
- Daily Staffing displays a prominent warning when groups remain unassigned or the busiest session needs more staff than are available.
- Accommodation navigation is replaced by School Notes. Every detected school has a notes box and a multi-staff assignment panel.


## v0.34
- Added Admin → Holidays monthly calendar.
- Head of Centre and Activities Manager can add and remove holidays.
- Added Team Leader login role with read-only holiday calendar access.
- Instructors cannot access the holiday calendar.


## v0.34 Team Leader access

- Team Leaders now open the full manager application.
- Team Leaders can use Home, Programme, Staffing, Arrivals, School Notes, Staff and Sign-off.
- Team Leaders can view the Holiday calendar but cannot add, edit or delete holidays.
- Head of Centre and Activities Manager retain full Holiday management.
- Supabase operational policies now recognise Team Leaders while holiday write policies remain restricted.


## v0.34 — My Sessions for operational roles

- Head of Centre, Activities Manager and Team Leader accounts now see a My Sessions panel on Home.
- The panel shows their own published activity and arrival duties while they remain inside the main management app.
- Duties are matched using the logged-in email address and refresh when the published rota changes.
