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
      and role = 'manager'
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
alter publication supabase_realtime add table public.rota_assignments;

-- IMPORTANT:
-- 1. Create your own account through the app.
-- 2. In Supabase Table Editor -> profiles, change your role from staff to manager.
-- 3. Staff create their accounts with the exact email entered on Staff Management.
