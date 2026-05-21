create extension if not exists pgcrypto;

create table if not exists public.pmo_records (
  id uuid primary key default gen_random_uuid(),
  type text not null check (type in ('projects','tasks','risks','reports','resources','schedules','files')),
  record jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists pmo_records_type_idx on public.pmo_records(type);
create index if not exists pmo_records_updated_at_idx on public.pmo_records(updated_at desc);

create or replace function public.set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists pmo_records_set_updated_at on public.pmo_records;
create trigger pmo_records_set_updated_at
before update on public.pmo_records
for each row
execute function public.set_updated_at();

alter table public.pmo_records enable row level security;

drop policy if exists "pmo_records_service_all" on public.pmo_records;
create policy "pmo_records_service_all"
on public.pmo_records
for all
using (true)
with check (true);
