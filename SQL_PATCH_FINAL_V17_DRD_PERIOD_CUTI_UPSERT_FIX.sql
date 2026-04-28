begin;

-- FINAL V17 - Fix Periode Cuti error:
-- there is no unique or exclusion constraint matching the ON CONFLICT specification
--
-- Patch ini aman untuk database lama. Ia merapikan data periode cuti per driver + end masa dinas
-- dan membuat unique index non-partial supaya database tetap rapi. Aplikasi v17 juga sudah tidak
-- bergantung pada ON CONFLICT, jadi error tidak muncul lagi meski schema cache Supabase lambat reload.

alter table if exists public.drd_induction_periods add column if not exists masa_dinas_end_date date;
alter table if exists public.drd_induction_periods add column if not exists updated_at timestamptz default now();

update public.drd_induction_periods p
set masa_dinas_end_date = d.end_masa_dinas
from public.drivers d
where p.driver_id = d.id
  and p.masa_dinas_end_date is null;

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

drop index if exists public.uq_drd_induction_period_driver_masa_dinas;

create unique index if not exists uq_drd_induction_period_driver_masa_dinas
on public.drd_induction_periods(driver_id, masa_dinas_end_date);

notify pgrst, 'reload schema';

commit;
