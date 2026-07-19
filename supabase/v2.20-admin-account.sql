-- Adventure Centre Manager v2.20
-- Run this once in Supabase SQL Editor before assigning an Admin account.

alter table public.profiles drop constraint if exists profiles_role_check;
alter table public.profiles add constraint profiles_role_check
  check (role in ('manager','staff','centreManager','activityManager','teamLeader','admin'));

-- Let managers find an existing account by email when granting Admin access.
drop policy if exists "Managers read profiles" on public.profiles;
create policy "Managers read profiles"
on public.profiles for select
to authenticated
using (public.is_manager());
