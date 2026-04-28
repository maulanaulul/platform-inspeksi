begin;

-- =============================================================
-- FINAL V16 - DRD cleanup: Master Driver, Periode Cuti, Results split
-- =============================================================

-- Master Driver tetap menyimpan masa dinas. Cuti dan onsite dipusatkan di drd_induction_periods.
alter table if exists public.drivers add column if not exists mulai_dinas date;
alter table if exists public.drivers add column if not exists end_masa_dinas date;
alter table if exists public.drivers add column if not exists updated_at timestamptz default now();

-- Periode cuti/onsite dibuat per periode end masa dinas, supaya ketika masa dinas baru diset,
-- outstanding input periode cuti dapat muncul lagi untuk periode berikutnya.
alter table if exists public.drd_induction_periods add column if not exists masa_dinas_end_date date;
alter table if exists public.drd_induction_periods add column if not exists updated_at timestamptz default now();

-- Isi masa_dinas_end_date untuk data lama jika masih kosong.
update public.drd_induction_periods p
set masa_dinas_end_date = d.end_masa_dinas
from public.drivers d
where p.driver_id = d.id
  and p.masa_dinas_end_date is null;

-- Bersihkan potensi duplikasi sebelum unique index dibuat.
with ranked as (
  select id,
         row_number() over (
           partition by driver_id, masa_dinas_end_date
           order by coalesce(updated_at, created_at) desc, created_at desc
         ) as rn
  from public.drd_induction_periods
  where masa_dinas_end_date is not null
)
delete from public.drd_induction_periods p
using ranked r
where p.id = r.id
  and r.rn > 1;

create unique index if not exists uq_drd_induction_period_driver_masa_dinas
on public.drd_induction_periods(driver_id, masa_dinas_end_date)
where masa_dinas_end_date is not null;

-- Pastikan attempt induksi tetap bisa mengacu ke periode cuti.
alter table if exists public.drd_attempts add column if not exists test_type text default 'DRD';
alter table if exists public.drd_attempts add column if not exists induction_period_id uuid;
update public.drd_attempts set test_type = coalesce(nullif(test_type,''), 'DRD');

notify pgrst, 'reload schema';
commit;
