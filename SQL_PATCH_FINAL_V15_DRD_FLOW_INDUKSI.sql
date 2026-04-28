begin;

-- =============================================================
-- FINAL V15 - DRD Driver Flow Simplification + Induksi Driver
-- =============================================================

-- Master driver masa dinas dan cuti/onsite
alter table if exists public.drivers add column if not exists mulai_dinas date;
alter table if exists public.drivers add column if not exists end_masa_dinas date;
alter table if exists public.drivers add column if not exists cuti_start_date date;
alter table if exists public.drivers add column if not exists onsite_date date;
alter table if exists public.drivers add column if not exists updated_at timestamptz default now();

-- Bank soal sekarang langsung memakai kategori, tanpa paket soal wajib.
alter table if exists public.drd_questions add column if not exists category text default 'DRD';
alter table if exists public.drd_questions add column if not exists status text default 'Aktif';
alter table if exists public.drd_questions alter column package_id drop not null;
update public.drd_questions set category = coalesce(nullif(category,''), 'DRD'), status = coalesce(status, 'Aktif');

-- Hasil test menampung DRD dan Induksi Driver.
alter table if exists public.drd_attempts add column if not exists test_type text default 'DRD';
alter table if exists public.drd_attempts add column if not exists induction_period_id uuid;
alter table if exists public.drd_attempts alter column package_id drop not null;
alter table if exists public.drd_attempts alter column assignment_id drop not null;
update public.drd_attempts set test_type = coalesce(nullif(test_type,''), 'DRD');

-- Video induksi.
create table if not exists public.drd_induction_videos (
  id uuid primary key default gen_random_uuid(),
  title text not null default 'Video Induksi Driver',
  video_url text not null,
  status text default 'Aktif',
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

-- Periode cuti/onsite untuk induksi.
create table if not exists public.drd_induction_periods (
  id uuid primary key default gen_random_uuid(),
  driver_id uuid references public.drivers(id) on delete cascade,
  site_id uuid references public.sites(id) on delete restrict,
  cuti_start_date date,
  onsite_date date,
  status text default 'Open',
  completed_at timestamptz,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

alter table if exists public.drd_induction_periods add column if not exists driver_id uuid references public.drivers(id) on delete cascade;
alter table if exists public.drd_induction_periods add column if not exists site_id uuid references public.sites(id) on delete restrict;
alter table if exists public.drd_induction_periods add column if not exists cuti_start_date date;
alter table if exists public.drd_induction_periods add column if not exists onsite_date date;
alter table if exists public.drd_induction_periods add column if not exists status text default 'Open';
alter table if exists public.drd_induction_periods add column if not exists completed_at timestamptz;
alter table if exists public.drd_induction_periods add column if not exists updated_at timestamptz default now();

create unique index if not exists uq_drd_induction_period_driver_onsite
on public.drd_induction_periods(driver_id, onsite_date);

-- Add FK to attempts after periods exists.
do $$
begin
  if not exists (
    select 1 from information_schema.table_constraints
    where table_schema='public' and table_name='drd_attempts' and constraint_name='drd_attempts_induction_period_id_fkey'
  ) then
    alter table public.drd_attempts
    add constraint drd_attempts_induction_period_id_fkey
    foreign key (induction_period_id) references public.drd_induction_periods(id) on delete set null;
  end if;
exception when others then
  null;
end $$;

-- Seed minimal soal agar menu tidak kosong di awal.
insert into public.drd_questions (package_id, category, question_text, option_a, option_b, option_c, option_d, correct_answer, status)
select null, 'DRD', 'Apa tindakan yang benar jika driver mulai mengantuk saat bekerja?', 'Tetap melanjutkan pekerjaan', 'Berhenti di area aman dan istirahat', 'Menambah kecepatan agar cepat selesai', 'Mengabaikan kondisi tubuh', 'B', 'Aktif'
where not exists (select 1 from public.drd_questions where category='DRD');

insert into public.drd_questions (package_id, category, question_text, option_a, option_b, option_c, option_d, correct_answer, status)
select null, 'Induksi Driver', 'Apa yang wajib dilakukan driver setelah kembali onsite dari cuti?', 'Langsung bekerja tanpa briefing', 'Mengikuti induksi sesuai periode datang cuti', 'Menunggu sampai dipanggil tanpa laporan', 'Mengabaikan prosedur site', 'B', 'Aktif'
where not exists (select 1 from public.drd_questions where category='Induksi Driver');

notify pgrst, 'reload schema';
commit;
