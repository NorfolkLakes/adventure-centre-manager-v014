# Adventure Centre Manager v0.14

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
