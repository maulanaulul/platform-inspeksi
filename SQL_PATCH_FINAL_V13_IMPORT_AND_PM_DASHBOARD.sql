-- Patch Final v13: sinkronisasi tahun plan berdasarkan created_at.
-- Perubahan template upload target_code/vendor_code berada di frontend, tidak butuh perubahan tabel.

begin;

update public.inspection_plans
set tahun = extract(year from created_at)::int
where created_at is not null
  and tahun is distinct from extract(year from created_at)::int;

-- Pastikan kategori PM Check tetap tersedia jika patch v12 belum pernah dijalankan.
alter table if exists public.inspection_parameters
  drop constraint if exists inspection_parameters_category_check;

alter table if exists public.inspection_parameters
  add constraint inspection_parameters_category_check
  check (category in ('Inspeksi Unit','Inspeksi Kelayakan Parkiran','PM Check'));

alter table if exists public.inspection_plans
  drop constraint if exists inspection_plans_category_check;

alter table if exists public.inspection_plans
  add constraint inspection_plans_category_check
  check (category is null or category in ('Inspeksi Unit','Inspeksi Kelayakan Parkiran','PM Check'));

insert into public.inspection_parameters (category, parameter_name, description, severity, status)
select
  'PM Check',
  'Apakah PM check telah dilakukan pada tanggal xxx (sesuai yang di planning)',
  'Pilih Aman jika PM check sudah dilakukan sesuai tanggal planning. Pilih Tidak Aman jika belum dilakukan / tidak sesuai jadwal.',
  'High',
  'Aktif'
where not exists (
  select 1 from public.inspection_parameters
  where category = 'PM Check'
    and parameter_name = 'Apakah PM check telah dilakukan pada tanggal xxx (sesuai yang di planning)'
);

notify pgrst, 'reload schema';

commit;
