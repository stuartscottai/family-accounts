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
