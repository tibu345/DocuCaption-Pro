create extension if not exists "pgcrypto";

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  email text not null,
  full_name text,
  plan text not null default 'free' check (plan in ('free', 'trial', 'annual', 'basic', 'pro')),
  subscription_status text not null default 'inactive' check (subscription_status in ('inactive', 'active', 'past_due', 'canceled')),
  documents_used_this_month integer not null default 0 check (documents_used_this_month >= 0),
  exports_used_this_month integer not null default 0 check (exports_used_this_month >= 0),
  caption_credits_remaining integer not null default 5000 check (caption_credits_remaining >= 0),
  billing_period_start timestamptz not null default now(),
  billing_period_end timestamptz not null default (now() + interval '1 month'),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.documents (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  file_name text not null,
  asset_count integer not null default 0 check (asset_count >= 0),
  status text not null default 'processed' check (status in ('uploaded', 'processed', 'exported', 'failed')),
  storage_bucket text,
  storage_path text,
  processed_storage_path text,
  exported_at timestamptz,
  deleted_at timestamptz,
  created_at timestamptz not null default now()
);

alter table public.documents
  add column if not exists storage_bucket text,
  add column if not exists storage_path text,
  add column if not exists processed_storage_path text,
  add column if not exists deleted_at timestamptz;

create table if not exists public.usage_logs (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  action text not null check (action in ('document_processed', 'export_created', 'captions_generated', 'captions_fallback', 'caption_denied')),
  amount integer not null default 1 check (amount >= 0),
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create or replace function public.handle_new_auth_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (
    id,
    email,
    full_name,
    plan,
    subscription_status,
    documents_used_this_month,
    exports_used_this_month,
    caption_credits_remaining,
    billing_period_start,
    billing_period_end
  )
  values (
    new.id,
    coalesce(new.email, ''),
    coalesce(new.raw_user_meta_data->>'full_name', new.raw_user_meta_data->>'name'),
    'free',
    'inactive',
    0,
    0,
    5000,
    now(),
    now() + interval '1 month'
  )
  on conflict (id) do update
  set
    email = excluded.email,
    full_name = coalesce(public.profiles.full_name, excluded.full_name),
    updated_at = now();

  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_auth_user();

insert into public.profiles (
  id,
  email,
  full_name,
  plan,
  subscription_status,
  documents_used_this_month,
  exports_used_this_month,
  caption_credits_remaining,
  billing_period_start,
  billing_period_end
)
select
  users.id,
  coalesce(users.email, ''),
  coalesce(users.raw_user_meta_data->>'full_name', users.raw_user_meta_data->>'name'),
  'free',
  'inactive',
  0,
  0,
  5000,
  now(),
  now() + interval '1 month'
from auth.users users
on conflict (id) do nothing;

insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values (
  'docucaption-documents',
  'docucaption-documents',
  false,
  52428800,
  array[
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/octet-stream'
  ]
)
on conflict (id) do update
set
  public = false,
  file_size_limit = excluded.file_size_limit,
  allowed_mime_types = excluded.allowed_mime_types;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'documents_status_check'
      and conrelid = 'public.documents'::regclass
  ) then
    alter table public.documents
      add constraint documents_status_check
      check (status in ('uploaded', 'processed', 'exported', 'failed'));
  end if;
end $$;

create index if not exists documents_user_id_idx on public.documents(user_id);
create index if not exists documents_status_idx on public.documents(status);
create index if not exists documents_created_at_idx on public.documents(created_at);
create index if not exists usage_logs_user_id_idx on public.usage_logs(user_id);
create index if not exists usage_logs_created_at_idx on public.usage_logs(created_at);
create index if not exists profiles_plan_idx on public.profiles(plan);
create index if not exists profiles_subscription_status_idx on public.profiles(subscription_status);

alter table public.profiles
  alter column plan set default 'free',
  alter column caption_credits_remaining set default 5000;

update public.profiles
set plan = 'free',
    caption_credits_remaining = greatest(caption_credits_remaining, 5000),
    updated_at = now()
where plan in ('trial', 'annual', 'free', 'basic', 'pro')
  and (plan <> 'free' or caption_credits_remaining < 5000);

alter table public.profiles enable row level security;
alter table public.documents enable row level security;
alter table public.usage_logs enable row level security;

drop policy if exists "profiles_select_own" on public.profiles;
create policy "profiles_select_own"
  on public.profiles
  for select
  to authenticated
  using (auth.uid() = id);

drop policy if exists "documents_select_own" on public.documents;
create policy "documents_select_own"
  on public.documents
  for select
  to authenticated
  using (auth.uid() = user_id);

drop policy if exists "usage_logs_select_own" on public.usage_logs;
create policy "usage_logs_select_own"
  on public.usage_logs
  for select
  to authenticated
  using (auth.uid() = user_id);

drop policy if exists "docucaption_storage_select_own" on storage.objects;
create policy "docucaption_storage_select_own"
  on storage.objects
  for select
  to authenticated
  using (
    bucket_id = 'docucaption-documents'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

drop policy if exists "docucaption_storage_insert_own" on storage.objects;
create policy "docucaption_storage_insert_own"
  on storage.objects
  for insert
  to authenticated
  with check (
    bucket_id = 'docucaption-documents'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

drop policy if exists "docucaption_storage_update_own" on storage.objects;
create policy "docucaption_storage_update_own"
  on storage.objects
  for update
  to authenticated
  using (
    bucket_id = 'docucaption-documents'
    and (storage.foldername(name))[1] = auth.uid()::text
  )
  with check (
    bucket_id = 'docucaption-documents'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

drop policy if exists "docucaption_storage_delete_own" on storage.objects;
create policy "docucaption_storage_delete_own"
  on storage.objects
  for delete
  to authenticated
  using (
    bucket_id = 'docucaption-documents'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

-- Intentionally do not add client insert/update/delete policies.
-- The Express backend uses SUPABASE_SERVICE_ROLE_KEY for metered writes.

create or replace function public.reserve_caption_credits(target_user_id uuid, credit_amount integer)
returns setof public.profiles
language sql
security definer
set search_path = public
as $$
  update public.profiles
  set
    caption_credits_remaining = caption_credits_remaining - credit_amount,
    updated_at = now()
  where
    id = target_user_id
    and credit_amount > 0
    and caption_credits_remaining >= credit_amount
  returning *;
$$;

create or replace function public.refund_caption_credits(target_user_id uuid, credit_amount integer)
returns setof public.profiles
language sql
security definer
set search_path = public
as $$
  update public.profiles
  set
    caption_credits_remaining = caption_credits_remaining + greatest(credit_amount, 0),
    updated_at = now()
  where id = target_user_id
  returning *;
$$;

create or replace function public.consume_document_quota(target_user_id uuid, monthly_limit integer)
returns setof public.profiles
language sql
security definer
set search_path = public
as $$
  update public.profiles
  set
    documents_used_this_month = documents_used_this_month + 1,
    updated_at = now()
  where
    id = target_user_id
    and documents_used_this_month < monthly_limit
  returning *;
$$;

create or replace function public.consume_export_quota(target_user_id uuid, monthly_limit integer)
returns setof public.profiles
language sql
security definer
set search_path = public
as $$
  update public.profiles
  set
    exports_used_this_month = exports_used_this_month + 1,
    updated_at = now()
  where
    id = target_user_id
    and exports_used_this_month < monthly_limit
  returning *;
$$;

revoke all on function public.reserve_caption_credits(uuid, integer) from public, anon, authenticated;
revoke all on function public.refund_caption_credits(uuid, integer) from public, anon, authenticated;
revoke all on function public.consume_document_quota(uuid, integer) from public, anon, authenticated;
revoke all on function public.consume_export_quota(uuid, integer) from public, anon, authenticated;

grant execute on function public.reserve_caption_credits(uuid, integer) to service_role;
grant execute on function public.refund_caption_credits(uuid, integer) to service_role;
grant execute on function public.consume_document_quota(uuid, integer) to service_role;
grant execute on function public.consume_export_quota(uuid, integer) to service_role;
