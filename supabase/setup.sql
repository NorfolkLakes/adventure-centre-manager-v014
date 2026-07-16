-- Adventure Centre Manager v0.13
-- Run this entire file in Supabase: SQL Editor -> New query -> Run.

create extension if not exists pgcrypto;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  email text not null,
  display_name text,
  role text not null default 'staff'
    check (role in ('manager', 'staff')),
  created_at timestamptz not null default now()
);

create table if not exists public.rota_assignments (
  id uuid primary key default gen_random_uuid(),
  programme_name text not null,
  day text not null,
  session text not null,
  activity_code text not null,
  activity_name text not null,
  group_numbers integer[] not null default '{}',
  duty_type text not null default 'activity',
  staff_email text not null,
  staff_name text not null,
  school_name text,
  published_at timestamptz not null default now()
);

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer set search_path = ''
as $$
begin
  insert into public.profiles (id, email, display_name, role)
  values (
    new.id,
    lower(coalesce(new.email, '')),
    coalesce(new.raw_user_meta_data ->> 'display_name', ''),
    'staff'
  )
  on conflict (id) do nothing;
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute procedure public.handle_new_user();

create or replace function public.is_manager()
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select exists (
    select 1
    from public.profiles
    where id = auth.uid()
      and role in ('manager','centreManager','activityManager','teamLeader')
  );
$$;

alter table public.profiles enable row level security;
alter table public.rota_assignments enable row level security;

drop policy if exists "Users read own profile" on public.profiles;
create policy "Users read own profile"
on public.profiles for select
to authenticated
using (id = auth.uid() or public.is_manager());

drop policy if exists "Managers update profiles" on public.profiles;
create policy "Managers update profiles"
on public.profiles for update
to authenticated
using (public.is_manager())
with check (public.is_manager());

drop policy if exists "Staff read own rota" on public.rota_assignments;
create policy "Staff read own rota"
on public.rota_assignments for select
to authenticated
using (
  lower(staff_email) = lower(coalesce(auth.jwt() ->> 'email', ''))
  or public.is_manager()
);

drop policy if exists "Managers insert rota" on public.rota_assignments;
create policy "Managers insert rota"
on public.rota_assignments for insert
to authenticated
with check (public.is_manager());

drop policy if exists "Managers update rota" on public.rota_assignments;
create policy "Managers update rota"
on public.rota_assignments for update
to authenticated
using (public.is_manager())
with check (public.is_manager());

drop policy if exists "Managers delete rota" on public.rota_assignments;
create policy "Managers delete rota"
on public.rota_assignments for delete
to authenticated
using (public.is_manager());

-- Realtime support for staff rota updates.
do $$ begin
  alter publication supabase_realtime add table public.rota_assignments;
exception when duplicate_object then null; end $$;

-- IMPORTANT:
-- 1. Create your own account through the app.
-- 2. In Supabase Table Editor -> profiles, change your role from staff to manager.
-- 3. Staff create their accounts with the exact email entered on Staff Management.


-- v0.18 staff self-service availability
create table if not exists public.staff_availability (
  id uuid primary key default gen_random_uuid(),
  staff_email text not null,
  day text not null,
  status text not null check (status in ('available','holiday','sick')),
  updated_at timestamptz not null default now(),
  unique (staff_email, day)
);
alter table public.staff_availability enable row level security;
drop policy if exists "Staff manage own availability" on public.staff_availability;
drop policy if exists "Managers manage availability" on public.staff_availability;
create policy "Managers manage availability" on public.staff_availability for all to authenticated
using (public.is_manager())
with check (public.is_manager());
drop policy if exists "Managers read availability" on public.staff_availability;
create policy "Managers read availability" on public.staff_availability for select to authenticated using (public.is_manager() or lower(staff_email)=lower(coalesce(auth.jwt()->>'email','')));
do $$ begin
  alter publication supabase_realtime add table public.staff_availability;
exception when duplicate_object then null; end $$;


-- v0.22 arrival and accommodation duty details
alter table public.rota_assignments add column if not exists building_name text;
alter table public.rota_assignments add column if not exists party_leader_name text;
alter table public.rota_assignments add column if not exists arrival_time text;
alter table public.rota_assignments add column if not exists departure_day text;
alter table public.rota_assignments add column if not exists departure_time text;

grant usage on schema public to authenticated;
grant select, insert, update, delete on public.rota_assignments to authenticated;
grant select, insert, update, delete on public.staff_availability to authenticated;
grant usage, select on all sequences in schema public to authenticated;


-- v0.32 holiday calendar and role permissions
alter table public.profiles drop constraint if exists profiles_role_check;
alter table public.profiles add constraint profiles_role_check
  check (role in ('manager','staff','centreManager','activityManager','teamLeader'));

create or replace function public.can_manage_holidays()
returns boolean language sql stable security definer set search_path = '' as $$
  select exists (select 1 from public.profiles where id = auth.uid() and role in ('manager','centreManager','activityManager'));
$$;

create or replace function public.can_view_holidays()
returns boolean language sql stable security definer set search_path = '' as $$
  select exists (select 1 from public.profiles where id = auth.uid() and role in ('manager','centreManager','activityManager','teamLeader'));
$$;

create table if not exists public.staff_holidays (
  id uuid primary key default gen_random_uuid(),
  staff_email text not null default '',
  staff_name text not null,
  start_date date not null,
  end_date date not null,
  note text,
  created_at timestamptz not null default now(),
  created_by uuid references auth.users(id),
  check (end_date >= start_date)
);
alter table public.staff_holidays enable row level security;
drop policy if exists "Holiday viewers read" on public.staff_holidays;
create policy "Holiday viewers read" on public.staff_holidays for select to authenticated using (public.can_view_holidays());
drop policy if exists "Holiday managers insert" on public.staff_holidays;
create policy "Holiday managers insert" on public.staff_holidays for insert to authenticated with check (public.can_manage_holidays());
drop policy if exists "Holiday managers update" on public.staff_holidays;
create policy "Holiday managers update" on public.staff_holidays for update to authenticated using (public.can_manage_holidays()) with check (public.can_manage_holidays());
drop policy if exists "Holiday managers delete" on public.staff_holidays;
create policy "Holiday managers delete" on public.staff_holidays for delete to authenticated using (public.can_manage_holidays());
grant select, insert, update, delete on public.staff_holidays to authenticated;
do $$ begin alter publication supabase_realtime add table public.staff_holidays; exception when duplicate_object then null; end $$;


-- v0.38 shared live application state
create table if not exists public.app_live_state (
  id text primary key,
  state jsonb not null default '{}'::jsonb,
  updated_by_name text not null default '',
  updated_by_email text not null default '',
  section text not null default 'dashboard',
  updated_at timestamptz not null default now()
);
alter table public.app_live_state enable row level security;
drop policy if exists "Operational users read live state" on public.app_live_state;
create policy "Operational users read live state" on public.app_live_state for select to authenticated using (public.is_manager());
drop policy if exists "Operational users insert live state" on public.app_live_state;
create policy "Operational users insert live state" on public.app_live_state for insert to authenticated with check (public.is_manager());
drop policy if exists "Operational users update live state" on public.app_live_state;
create policy "Operational users update live state" on public.app_live_state for update to authenticated using (public.is_manager()) with check (public.is_manager());
grant select, insert, update on public.app_live_state to authenticated;
do $$ begin alter publication supabase_realtime add table public.app_live_state; exception when duplicate_object then null; end $$;
