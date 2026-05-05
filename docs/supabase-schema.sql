-- Family Accounts Supabase reference setup
-- Last reviewed from screenshots: 2026-05-05
--
-- This file documents the database structure this app expects.
-- It is a reference/checklist script. Do not run it blindly on a live
-- project without first checking whether the tables and policies already
-- exist, especially if there is real data in Supabase.

-- Required extension for auth.uid().
-- Supabase normally already provides this through the auth system.

-- ============================================================================
-- Main account table
-- Stores the global app file:
--   - people
--   - categories
--   - a months object used as a backup/legacy copy
-- ============================================================================

create table if not exists public.family_account_data (
  user_id uuid primary key references auth.users(id) on delete cascade,
  data jsonb not null default '{"people": [], "categories": [], "months": {}}'::jsonb,
  updated_at timestamptz not null default now()
);

alter table public.family_account_data enable row level security;

-- Recommended simple policy.
-- In plain English: a signed-in user can only read/write their own row.
drop policy if exists "Users can manage their own data" on public.family_account_data;

create policy "Users can manage their own data"
on public.family_account_data
for all
to public
using (auth.uid() = user_id)
with check (auth.uid() = user_id);

-- Screenshot state confirmed on 2026-05-05:
-- The live Supabase table had these policies:
--
--   1. Users can insert own data
--      command: insert
--      with check (auth.uid() = user_id)
--
--   2. Users can manage their globals
--      command: all
--      using (auth.uid() = user_id)
--      with check (auth.uid() = user_id)
--
--   3. Users can read own data
--      command: select
--      using (auth.uid() = user_id)
--
--   4. Users can update own data
--      command: update
--      using (auth.uid() = user_id)
--      with check (auth.uid() = user_id)
--
-- This is safe. It is slightly redundant because "Users can manage their
-- globals" already covers all commands, but the extra specific insert/select/
-- update policies do not weaken security because they use the same user_id
-- check.

-- ============================================================================
-- Monthly data table
-- Stores each month separately so one month's expenses can be loaded/saved
-- without rewriting the whole account file.
-- ============================================================================

create table if not exists public.family_account_months (
  user_id uuid not null references auth.users(id) on delete cascade,
  month text not null,
  data jsonb not null default '{"expenses": []}'::jsonb,
  updated_at timestamptz not null default now(),
  primary key (user_id, month)
);

alter table public.family_account_months enable row level security;

-- Recommended simple policy.
-- In plain English: a signed-in user can only read/write their own month rows.
drop policy if exists "Users can manage their months" on public.family_account_months;

create policy "Users can manage their months"
on public.family_account_months
for all
to public
using (auth.uid() = user_id)
with check (auth.uid() = user_id);

-- Screenshot state seen on 2026-05-05:
-- family_account_months had one ALL policy named:
--   "Users can manage their months"
--
-- The visible policy text was:
--   using (auth.uid() = user_id)
--   with check (auth.uid() = user_id)
--
-- That is the expected policy.

-- ============================================================================
-- Helpful checks
-- Run these in Supabase SQL Editor if you need to inspect setup.
-- ============================================================================

-- Check columns:
select
  table_name,
  column_name,
  data_type,
  is_nullable
from information_schema.columns
where table_schema = 'public'
  and table_name in ('family_account_data', 'family_account_months')
order by table_name, ordinal_position;

-- Check primary/unique constraints.
-- family_account_months must have a primary key or unique constraint on:
--   user_id, month
select
  tc.table_name,
  tc.constraint_name,
  tc.constraint_type,
  string_agg(kcu.column_name, ', ' order by kcu.ordinal_position) as columns
from information_schema.table_constraints tc
join information_schema.key_column_usage kcu
  on tc.constraint_name = kcu.constraint_name
 and tc.table_schema = kcu.table_schema
where tc.table_schema = 'public'
  and tc.table_name in ('family_account_data', 'family_account_months')
  and tc.constraint_type in ('PRIMARY KEY', 'UNIQUE')
group by tc.table_name, tc.constraint_name, tc.constraint_type
order by tc.table_name, tc.constraint_name;

-- Check RLS policies.
select
  schemaname,
  tablename,
  policyname,
  permissive,
  roles,
  cmd,
  qual as using_expression,
  with_check as check_expression
from pg_policies
where schemaname = 'public'
  and tablename in ('family_account_data', 'family_account_months')
order by tablename, policyname;

-- Check whether monthly rows contain expenses.
-- This does not show private details from the JSON, just counts.
select
  month,
  jsonb_array_length(coalesce(data->'expenses', '[]'::jsonb)) as expense_count,
  updated_at
from public.family_account_months
order by month desc;

-- ============================================================================
-- Backup history table
-- This table is optional for the app to run, but strongly recommended.
-- It keeps a history row whenever account/month data is inserted, updated,
-- or deleted. This is the recovery layer for accidental overwrites.
-- ============================================================================

create table if not exists public.family_account_backups (
  id bigserial primary key,
  user_id uuid not null references auth.users(id) on delete cascade,
  source_table text not null,
  source_month text,
  operation text not null,
  old_data jsonb,
  new_data jsonb,
  created_at timestamptz not null default now()
);

create index if not exists family_account_backups_user_created_idx
on public.family_account_backups (user_id, created_at desc);

create index if not exists family_account_backups_month_idx
on public.family_account_backups (user_id, source_month, created_at desc);

alter table public.family_account_backups enable row level security;

drop policy if exists "Users can read their own backups" on public.family_account_backups;

create policy "Users can read their own backups"
on public.family_account_backups
for select
to public
using (auth.uid() = user_id);

create or replace function public.backup_family_account_change()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  backup_user_id uuid;
  backup_month text;
begin
  backup_user_id := coalesce(new.user_id, old.user_id);

  if tg_table_name = 'family_account_months' then
    backup_month := coalesce(new.month, old.month);
  else
    backup_month := null;
  end if;

  insert into public.family_account_backups (
    user_id,
    source_table,
    source_month,
    operation,
    old_data,
    new_data
  )
  values (
    backup_user_id,
    tg_table_name,
    backup_month,
    tg_op,
    case when tg_op in ('UPDATE', 'DELETE') then to_jsonb(old) else null end,
    case when tg_op in ('INSERT', 'UPDATE') then to_jsonb(new) else null end
  );

  if tg_op = 'DELETE' then
    return old;
  end if;

  return new;
end;
$$;

drop trigger if exists backup_family_account_data_insert on public.family_account_data;
drop trigger if exists backup_family_account_data_update on public.family_account_data;
drop trigger if exists backup_family_account_data_delete on public.family_account_data;

create trigger backup_family_account_data_insert
after insert on public.family_account_data
for each row
execute function public.backup_family_account_change();

create trigger backup_family_account_data_update
after update on public.family_account_data
for each row
when (old.data is distinct from new.data)
execute function public.backup_family_account_change();

create trigger backup_family_account_data_delete
after delete on public.family_account_data
for each row
execute function public.backup_family_account_change();

drop trigger if exists backup_family_account_months_insert on public.family_account_months;
drop trigger if exists backup_family_account_months_update on public.family_account_months;
drop trigger if exists backup_family_account_months_delete on public.family_account_months;

create trigger backup_family_account_months_insert
after insert on public.family_account_months
for each row
execute function public.backup_family_account_change();

create trigger backup_family_account_months_update
after update on public.family_account_months
for each row
when (old.data is distinct from new.data)
execute function public.backup_family_account_change();

create trigger backup_family_account_months_delete
after delete on public.family_account_months
for each row
execute function public.backup_family_account_change();
