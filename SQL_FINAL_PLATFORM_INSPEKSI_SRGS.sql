-- =============================================================
-- PLATFORM INSPEKSI SRGS - FINAL ALL-IN-ONE
-- Gabungan 3 aplikasi: Sidak Fatigue, DRD Driver, Inspeksi
-- Jalankan sekali di Supabase SQL Editor.
-- Aman untuk database lama karena memakai IF NOT EXISTS / ADD COLUMN IF NOT EXISTS.
-- =============================================================

create extension if not exists pgcrypto;

-- Role enum compatibility jika database lama memakai enum app_role.
do $$
begin
  if exists (select 1 from pg_type where typname = 'app_role') then
    alter type app_role add value if not exists 'Platform Admin';
    alter type app_role add value if not exists 'App Admin';
    alter type app_role add value if not exists 'GL';
    alter type app_role add value if not exists 'Atasan Site';
    alter type app_role add value if not exists 'Driver';
    alter type app_role add value if not exists 'Viewer';
  end if;
end $$;

-- Plan status enum compatibility jika database lama memakai enum plan_status.
do $$
begin
  if exists (select 1 from pg_type where typname = 'plan_status') then
    alter type plan_status add value if not exists 'Planned';
    alter type plan_status add value if not exists 'Submitted';
    alter type plan_status add value if not exists 'In Review';
    alter type plan_status add value if not exists 'Approved';
    alter type plan_status add value if not exists 'Rejected';
    alter type plan_status add value if not exists 'Done';
  end if;
end $$;

-- =============================================================
-- CORE PLATFORM
-- =============================================================
create table if not exists public.applications (
  id uuid primary key default gen_random_uuid(),
  app_code text unique not null,
  app_name text not null,
  description text,
  status text default 'Aktif',
  created_at timestamptz default now()
);

create table if not exists public.sites (
  id uuid primary key default gen_random_uuid(),
  site_code text unique not null,
  site_name text not null,
  region text,
  status text default 'Aktif',
  created_at timestamptz default now()
);

create table if not exists public.users_profile (
  id uuid primary key default gen_random_uuid(),
  auth_user_id uuid unique,
  nama text,
  nrp text,
  email text unique not null,
  status text default 'Aktif',
  created_at timestamptz default now()
);

create table if not exists public.user_app_access (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references public.users_profile(id) on delete cascade,
  app_id uuid not null references public.applications(id) on delete cascade,
  role text not null,
  site_id uuid references public.sites(id) on delete set null,
  status text default 'Aktif',
  created_at timestamptz default now()
);

create unique index if not exists uq_user_app_access_active
on public.user_app_access (user_id, app_id, role, coalesce(site_id, '00000000-0000-0000-0000-000000000000'::uuid))
where status = 'Aktif';

insert into public.applications (app_code, app_name, description, status) values
('sidak_fatigue', 'Sidak Fatigue', 'Aplikasi sidak fatigue driver', 'Aktif'),
('drd_driver', 'DRD Driver', 'Aplikasi DRD dan test driver', 'Aktif'),
('inspeksi_unit', 'Inspeksi', 'Aplikasi inspeksi unit dan kelayakan parkiran', 'Aktif')
on conflict (app_code) do update set app_name = excluded.app_name, description = excluded.description, status = 'Aktif';

insert into public.sites (site_code, site_name, region, status) values
('ABKL','ABKL','Operation','Aktif'),('ARIA','ARIA','Operation','Aktif'),('ASMI','ASMI','Operation','Aktif'),
('BAYA','BAYA','Operation','Aktif'),('BBSO','BBSO','Operation','Aktif'),('BEKB','BEKB','Operation','Aktif'),
('BPOP','BPOP','Operation','Aktif'),('BRCB','BRCB','Operation','Aktif'),('BRCG','BRCG','Operation','Aktif'),
('BRCS','BRCS','Operation','Aktif'),('BTSJ','BTSJ','Operation','Aktif'),('HMNT','HMNT','Operation','Aktif'),
('INDO','INDO','Operation','Aktif'),('KIDE','KIDE','Operation','Aktif'),('KPCB','KPCB','Operation','Aktif'),
('KPCS','KPCS','Operation','Aktif'),('KPCT','KPCT','Operation','Aktif'),('MTBU','MTBU','Operation','Aktif'),
('NPRL','NPRL','Operation','Aktif'),('SMMS','SMMS','Operation','Aktif'),('TCMM','TCMM','Operation','Aktif'),
('TOPB','TOPB','Operation','Aktif'),('VIPO','VIPO','Operation','Aktif'),('JIEP','JIEP / Head Office','Head Office','Aktif')
on conflict (site_code) do update set site_name = excluded.site_name, region = excluded.region, status = 'Aktif';

-- =============================================================
-- SHARED MASTER: VENDOR & DRIVER
-- =============================================================
create table if not exists public.vendors (
  id uuid primary key default gen_random_uuid(),
  vendor_code text unique,
  vendor_name text not null,
  status text default 'Aktif',
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create or replace function public.trg_set_vendor_code()
returns trigger language plpgsql as $$
begin
  if new.id is null then new.id := gen_random_uuid(); end if;
  if nullif(trim(coalesce(new.vendor_code,'')),'') is null then
    new.vendor_code := 'VEN-' || upper(substr(new.id::text,1,8));
  else
    new.vendor_code := upper(trim(new.vendor_code));
  end if;
  return new;
end;
$$;
drop trigger if exists set_vendor_code on public.vendors;
create trigger set_vendor_code before insert on public.vendors for each row execute function public.trg_set_vendor_code();

update public.vendors set vendor_code = 'VEN-' || upper(substr(id::text,1,8)) where nullif(trim(coalesce(vendor_code,'')),'') is null;

create table if not exists public.drivers (
  id uuid primary key default gen_random_uuid(),
  site_id uuid references public.sites(id) on delete restrict,
  vendor_id uuid references public.vendors(id) on delete set null,
  nama_driver text not null,
  nrp_driver text unique not null,
  email text,
  status text default 'Aktif',
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

alter table if exists public.drivers add column if not exists site_id uuid references public.sites(id) on delete restrict;
alter table if exists public.drivers add column if not exists vendor_id uuid references public.vendors(id) on delete set null;
alter table if exists public.drivers add column if not exists email text;
alter table if exists public.drivers add column if not exists status text default 'Aktif';
alter table if exists public.drivers add column if not exists updated_at timestamptz default now();

-- =============================================================
-- SIDAK FATIGUE
-- =============================================================
create table if not exists public.fatigue_parameters (
  id uuid primary key default gen_random_uuid(),
  parameter_name text not null,
  description text,
  urutan int,
  status text default 'Aktif',
  created_at timestamptz default now()
);

-- Compatibility for databases where fatigue_parameters was created by an older package.
alter table if exists public.fatigue_parameters add column if not exists parameter_name text;
alter table if exists public.fatigue_parameters add column if not exists description text;
alter table if exists public.fatigue_parameters add column if not exists urutan int;
alter table if exists public.fatigue_parameters add column if not exists status text default 'Aktif';
alter table if exists public.fatigue_parameters add column if not exists created_at timestamptz default now();

update public.fatigue_parameters
set parameter_name = coalesce(parameter_name, 'Parameter Fatigue'),
    status = coalesce(status, 'Aktif'),
    created_at = coalesce(created_at, now())
where parameter_name is null or status is null or created_at is null;

insert into public.fatigue_parameters (parameter_name, description, urutan, status) values
('Kondisi mengantuk', 'Driver menunjukkan tanda mengantuk/fatigue.', 1, 'Aktif'),
('Kesiapan fisik', 'Kondisi fisik driver layak bekerja.', 2, 'Aktif'),
('Kepatuhan istirahat', 'Driver telah memenuhi waktu istirahat.', 3, 'Aktif')
on conflict do nothing;

create table if not exists public.fatigue_plans (
  id uuid primary key default gen_random_uuid(),
  site_id uuid references public.sites(id) on delete restrict,
  driver_id uuid references public.drivers(id) on delete cascade,
  bulan int,
  tahun int,
  status text default 'Planned',
  created_by uuid references public.users_profile(id) on delete set null,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

alter table if exists public.fatigue_plans add column if not exists site_id uuid references public.sites(id) on delete restrict;
alter table if exists public.fatigue_plans add column if not exists driver_id uuid references public.drivers(id) on delete cascade;
alter table if exists public.fatigue_plans add column if not exists bulan int;
alter table if exists public.fatigue_plans add column if not exists tahun int;
alter table if exists public.fatigue_plans add column if not exists status text default 'Planned';
alter table if exists public.fatigue_plans add column if not exists created_by uuid references public.users_profile(id) on delete set null;
alter table if exists public.fatigue_plans add column if not exists updated_at timestamptz default now();

do $$ begin
  if exists (select 1 from information_schema.columns where table_schema='public' and table_name='fatigue_plans' and column_name='object_id') then
    alter table public.fatigue_plans alter column object_id drop not null;
  end if;
end $$;

create table if not exists public.fatigue_inspections (
  id uuid primary key default gen_random_uuid(),
  plan_id uuid references public.fatigue_plans(id) on delete set null,
  site_id uuid references public.sites(id) on delete restrict,
  driver_id uuid references public.drivers(id) on delete restrict,
  tanggal_inspeksi date default current_date,
  status text default 'Submitted',
  catatan text,
  foto_inspeksi_url text,
  inspected_by uuid references public.users_profile(id) on delete set null,
  approved_by uuid references public.users_profile(id) on delete set null,
  approved_at timestamptz,
  approval_note text,
  created_at timestamptz default now()
);

create table if not exists public.fatigue_inspection_details (
  id uuid primary key default gen_random_uuid(),
  inspection_id uuid references public.fatigue_inspections(id) on delete cascade,
  parameter_id uuid references public.fatigue_parameters(id) on delete restrict,
  hasil text default 'Aman',
  note text,
  created_at timestamptz default now()
);

create table if not exists public.fatigue_outstandings (
  id uuid primary key default gen_random_uuid(),
  inspection_id uuid references public.fatigue_inspections(id) on delete set null,
  detail_id uuid references public.fatigue_inspection_details(id) on delete set null,
  site_id uuid references public.sites(id) on delete restrict,
  driver_id uuid references public.drivers(id) on delete restrict,
  parameter_id uuid references public.fatigue_parameters(id) on delete set null,
  description text,
  status text default 'Open',
  close_note text,
  evidence_photo_url text,
  close_requested_by uuid references public.users_profile(id) on delete set null,
  close_requested_at timestamptz,
  approved_by uuid references public.users_profile(id) on delete set null,
  approved_at timestamptz,
  approval_note text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

-- =============================================================
-- DRD DRIVER
-- =============================================================
create table if not exists public.drd_question_packages (
  id uuid primary key default gen_random_uuid(),
  package_name text not null,
  description text,
  question_count int default 10,
  status text default 'Aktif',
  created_at timestamptz default now()
);

create table if not exists public.drd_questions (
  id uuid primary key default gen_random_uuid(),
  package_id uuid references public.drd_question_packages(id) on delete cascade,
  question_text text not null,
  option_a text,
  option_b text,
  option_c text,
  option_d text,
  correct_answer text,
  category text,
  urutan int,
  status text default 'Aktif',
  created_at timestamptz default now()
);

create table if not exists public.drd_assignments (
  id uuid primary key default gen_random_uuid(),
  driver_id uuid references public.drivers(id) on delete cascade,
  package_id uuid references public.drd_question_packages(id) on delete restrict,
  site_id uuid references public.sites(id),
  assigned_by uuid references public.users_profile(id) on delete set null,
  due_date date,
  status text default 'Belum Test',
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists public.drd_attempts (
  id uuid primary key default gen_random_uuid(),
  assignment_id uuid references public.drd_assignments(id) on delete set null,
  driver_id uuid references public.drivers(id) on delete cascade,
  package_id uuid references public.drd_question_packages(id) on delete set null,
  submitted_at timestamptz,
  score int default 0,
  total_questions int default 0,
  correct_count int default 0,
  status text default 'Tidak Lulus',
  valid_until date,
  created_at timestamptz default now()
);

create table if not exists public.drd_answers (
  id uuid primary key default gen_random_uuid(),
  attempt_id uuid references public.drd_attempts(id) on delete cascade,
  question_id uuid references public.drd_questions(id) on delete set null,
  selected_answer text,
  is_correct boolean default false,
  created_at timestamptz default now()
);

alter table if exists public.drd_question_packages add column if not exists question_count int default 10;
alter table if exists public.drd_questions add column if not exists urutan int;
alter table if exists public.drd_assignments add column if not exists site_id uuid references public.sites(id);
alter table if exists public.drd_attempts add column if not exists package_id uuid references public.drd_question_packages(id);

insert into public.drd_question_packages(package_name, description, question_count, status) values
('Paket DRD Dasar', 'Paket awal untuk testing DRD Driver', 5, 'Aktif')
on conflict do nothing;

-- =============================================================
-- INSPEKSI UNIT & PARKIRAN
-- =============================================================
create table if not exists public.inspection_units (
  id uuid primary key default gen_random_uuid(),
  site_id uuid not null references public.sites(id) on delete restrict,
  unit_code text,
  unit_name text not null,
  unit_type text,
  location text,
  status text default 'Aktif',
  created_by uuid references public.users_profile(id) on delete set null,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);
create unique index if not exists uq_inspection_units_site_code on public.inspection_units(site_id, unit_code);

create table if not exists public.inspection_parkings (
  id uuid primary key default gen_random_uuid(),
  site_id uuid not null references public.sites(id) on delete restrict,
  parking_code text,
  parking_name text not null,
  location text,
  capacity int,
  status text default 'Aktif',
  created_by uuid references public.users_profile(id) on delete set null,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);
create unique index if not exists uq_inspection_parkings_site_code on public.inspection_parkings(site_id, parking_code);

create table if not exists public.inspection_parameters (
  id uuid primary key default gen_random_uuid(),
  category text not null,
  parameter_code text,
  parameter_name text not null,
  description text,
  severity text default 'Medium',
  status text default 'Aktif',
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);
create unique index if not exists uq_inspection_parameters_category_code on public.inspection_parameters(category, parameter_code);

alter table if exists public.inspection_parameters drop constraint if exists inspection_parameters_category_check;
alter table if exists public.inspection_parameters add constraint inspection_parameters_category_check check (category in ('Inspeksi Unit','Inspeksi Kelayakan Parkiran','PM Check'));

create table if not exists public.inspection_plans (
  id uuid primary key default gen_random_uuid(),
  site_id uuid not null references public.sites(id) on delete restrict,
  category text not null,
  target_type text not null,
  unit_id uuid references public.inspection_units(id) on delete restrict,
  parking_id uuid references public.inspection_parkings(id) on delete restrict,
  bulan int,
  tahun int,
  due_date date,
  status text default 'Planned',
  planned_by uuid references public.users_profile(id) on delete set null,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

alter table if exists public.inspection_plans add column if not exists site_id uuid references public.sites(id) on delete restrict;
alter table if exists public.inspection_plans add column if not exists category text;
alter table if exists public.inspection_plans add column if not exists target_type text;
alter table if exists public.inspection_plans add column if not exists unit_id uuid references public.inspection_units(id) on delete restrict;
alter table if exists public.inspection_plans add column if not exists parking_id uuid references public.inspection_parkings(id) on delete restrict;
alter table if exists public.inspection_plans add column if not exists bulan int;
alter table if exists public.inspection_plans add column if not exists tahun int;
alter table if exists public.inspection_plans add column if not exists due_date date;
alter table if exists public.inspection_plans add column if not exists status text default 'Planned';
alter table if exists public.inspection_plans add column if not exists planned_by uuid references public.users_profile(id) on delete set null;
alter table if exists public.inspection_plans add column if not exists updated_at timestamptz default now();
alter table if exists public.inspection_plans drop constraint if exists inspection_plans_category_check;
alter table if exists public.inspection_plans add constraint inspection_plans_category_check check (category is null or category in ('Inspeksi Unit','Inspeksi Kelayakan Parkiran','PM Check'));
alter table if exists public.inspection_plans drop constraint if exists inspection_plans_target_type_check;
alter table if exists public.inspection_plans add constraint inspection_plans_target_type_check check (target_type is null or target_type in ('unit','parkiran'));

do $$ begin
  if exists (select 1 from information_schema.columns where table_schema='public' and table_name='inspection_plans' and column_name='object_id') then
    alter table public.inspection_plans alter column object_id drop not null;
  end if;
  if exists (select 1 from information_schema.columns where table_schema='public' and table_name='inspection_plans' and column_name='object_type') then
    alter table public.inspection_plans alter column object_type drop not null;
  end if;
end $$;

create table if not exists public.inspection_records (
  id uuid primary key default gen_random_uuid(),
  plan_id uuid references public.inspection_plans(id) on delete cascade,
  site_id uuid references public.sites(id) on delete restrict,
  category text not null,
  inspector_id uuid references public.users_profile(id) on delete set null,
  inspected_at timestamptz default now(),
  result text default 'Aman',
  status text default 'Submitted',
  notes text,
  photo_url text,
  approved_by uuid references public.users_profile(id) on delete set null,
  approved_at timestamptz,
  rejection_reason text,
  created_at timestamptz default now()
);

create table if not exists public.inspection_answers (
  id uuid primary key default gen_random_uuid(),
  record_id uuid references public.inspection_records(id) on delete cascade,
  parameter_id uuid references public.inspection_parameters(id) on delete restrict,
  result text default 'Aman',
  notes text,
  created_at timestamptz default now()
);

create table if not exists public.inspection_findings (
  id uuid primary key default gen_random_uuid(),
  record_id uuid references public.inspection_records(id) on delete set null,
  plan_id uuid references public.inspection_plans(id) on delete set null,
  site_id uuid references public.sites(id) on delete restrict,
  category text not null,
  target_label text,
  parameter_id uuid references public.inspection_parameters(id) on delete set null,
  finding_description text,
  priority text default 'Medium',
  due_date date,
  status text default 'Open',
  created_by uuid references public.users_profile(id) on delete set null,
  close_requested_by uuid references public.users_profile(id) on delete set null,
  close_note text,
  close_photo_url text,
  close_requested_at timestamptz,
  close_approved_by uuid references public.users_profile(id) on delete set null,
  close_approved_at timestamptz,
  close_rejection_reason text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

-- Auto-generate unique codes for inspection unit, parking, and parameter.
create or replace function public.trg_set_inspection_unit_code()
returns trigger language plpgsql as $$
begin
  if new.id is null then new.id := gen_random_uuid(); end if;
  if nullif(trim(coalesce(new.unit_code,'')),'') is null then
    new.unit_code := 'UNIT-' || upper(substr(new.id::text,1,8));
  else new.unit_code := upper(trim(new.unit_code)); end if;
  return new;
end; $$;
drop trigger if exists set_inspection_unit_code on public.inspection_units;
create trigger set_inspection_unit_code before insert on public.inspection_units for each row execute function public.trg_set_inspection_unit_code();

create or replace function public.trg_set_inspection_parking_code()
returns trigger language plpgsql as $$
begin
  if new.id is null then new.id := gen_random_uuid(); end if;
  if nullif(trim(coalesce(new.parking_code,'')),'') is null then
    new.parking_code := 'PARK-' || upper(substr(new.id::text,1,8));
  else new.parking_code := upper(trim(new.parking_code)); end if;
  return new;
end; $$;
drop trigger if exists set_inspection_parking_code on public.inspection_parkings;
create trigger set_inspection_parking_code before insert on public.inspection_parkings for each row execute function public.trg_set_inspection_parking_code();

create or replace function public.trg_set_inspection_parameter_code()
returns trigger language plpgsql as $$
declare code_prefix text;
begin
  if new.id is null then new.id := gen_random_uuid(); end if;
  code_prefix := case when new.category = 'Inspeksi Kelayakan Parkiran' then 'PPARK-' when new.category = 'PM Check' then 'PPM-' else 'PUNIT-' end;
  if nullif(trim(coalesce(new.parameter_code,'')),'') is null then
    new.parameter_code := code_prefix || upper(substr(new.id::text,1,8));
  else new.parameter_code := upper(trim(new.parameter_code)); end if;
  return new;
end; $$;
drop trigger if exists set_inspection_parameter_code on public.inspection_parameters;
create trigger set_inspection_parameter_code before insert on public.inspection_parameters for each row execute function public.trg_set_inspection_parameter_code();

update public.inspection_units set unit_code = 'UNIT-' || upper(substr(id::text,1,8)) where nullif(trim(coalesce(unit_code,'')),'') is null;
update public.inspection_parkings set parking_code = 'PARK-' || upper(substr(id::text,1,8)) where nullif(trim(coalesce(parking_code,'')),'') is null;
update public.inspection_parameters set parameter_code = (case when category = 'Inspeksi Kelayakan Parkiran' then 'PPARK-' when category = 'PM Check' then 'PPM-' else 'PUNIT-' end) || upper(substr(id::text,1,8)) where nullif(trim(coalesce(parameter_code,'')),'') is null;

insert into public.inspection_parameters (category, parameter_name, description, severity, status) values
('Inspeksi Unit','Kondisi fisik unit','Cek kerusakan, kebocoran, retak, atau kondisi tidak normal pada unit.','High','Aktif'),
('Inspeksi Unit','Kelengkapan safety unit','Cek APAR, segitiga pengaman, P3K, lampu, alarm mundur, dan kelengkapan safety lainnya.','High','Aktif'),
('Inspeksi Kelayakan Parkiran','Rambu dan marka parkir','Cek ketersediaan rambu, marka, arah masuk/keluar, dan jalur aman.','High','Aktif'),
('Inspeksi Kelayakan Parkiran','Penerangan area parkir','Cek kecukupan lampu dan visibilitas area parkir.','High','Aktif'),
('PM Check','Apakah PM check telah dilakukan pada tanggal xxx (sesuai yang di planning)','Pilih Aman jika PM check sudah dilakukan sesuai tanggal planning. Pilih Tidak Aman jika belum dilakukan / tidak sesuai jadwal.','High','Aktif')
on conflict do nothing;

-- =============================================================
-- STORAGE BUCKETS + POLICIES
-- =============================================================
insert into storage.buckets (id, name, public) values
('inspection-photos','inspection-photos',true),
('evidence-photos','evidence-photos',true),
('drd-assets','drd-assets',true),
('inspection-unit-photos','inspection-unit-photos',true),
('inspection-close-evidence','inspection-close-evidence',true)
on conflict (id) do nothing;

drop policy if exists "SRGS all inspection photos" on storage.objects;
drop policy if exists "SRGS all evidence photos" on storage.objects;
drop policy if exists "SRGS all drd assets" on storage.objects;
drop policy if exists "SRGS all inspection unit photos" on storage.objects;
drop policy if exists "SRGS all inspection close evidence" on storage.objects;
create policy "SRGS all inspection photos" on storage.objects for all to authenticated using (bucket_id='inspection-photos') with check (bucket_id='inspection-photos');
create policy "SRGS all evidence photos" on storage.objects for all to authenticated using (bucket_id='evidence-photos') with check (bucket_id='evidence-photos');
create policy "SRGS all drd assets" on storage.objects for all to authenticated using (bucket_id='drd-assets') with check (bucket_id='drd-assets');
create policy "SRGS all inspection unit photos" on storage.objects for all to authenticated using (bucket_id='inspection-unit-photos') with check (bucket_id='inspection-unit-photos');
create policy "SRGS all inspection close evidence" on storage.objects for all to authenticated using (bucket_id='inspection-close-evidence') with check (bucket_id='inspection-close-evidence');

-- =============================================================
-- SIMPLE RLS POLICY
-- =============================================================
do $$
declare t text;
begin
  foreach t in array array[
    'applications','sites','users_profile','user_app_access','vendors','drivers',
    'fatigue_parameters','fatigue_plans','fatigue_inspections','fatigue_inspection_details','fatigue_outstandings',
    'drd_question_packages','drd_questions','drd_assignments','drd_attempts','drd_answers',
    'inspection_units','inspection_parkings','inspection_parameters','inspection_plans','inspection_records','inspection_answers','inspection_findings'
  ] loop
    execute format('alter table public.%I enable row level security', t);
    execute format('drop policy if exists "SRGS authenticated all %s" on public.%I', t, t);
    execute format('create policy "SRGS authenticated all %s" on public.%I for all to authenticated using (true) with check (true)', t, t);
  end loop;
end $$;

notify pgrst, 'reload schema';

select app_code, app_name, status from public.applications order by app_code;
