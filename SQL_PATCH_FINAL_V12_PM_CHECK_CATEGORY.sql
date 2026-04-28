-- Patch Final v12: tambah kategori PM Check pada aplikasi Inspeksi.
-- Jalankan di Supabase SQL Editor setelah source v12 digunakan.

begin;

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

create or replace function public.trg_set_inspection_parameter_code()
returns trigger language plpgsql as $$
declare code_prefix text;
begin
  if new.id is null then new.id := gen_random_uuid(); end if;
  code_prefix := case
    when new.category = 'Inspeksi Kelayakan Parkiran' then 'PPARK-'
    when new.category = 'PM Check' then 'PPM-'
    else 'PUNIT-'
  end;

  if nullif(trim(coalesce(new.parameter_code,'')),'') is null then
    new.parameter_code := code_prefix || upper(substr(new.id::text,1,8));
  else
    new.parameter_code := upper(trim(new.parameter_code));
  end if;

  return new;
end;
$$;

drop trigger if exists set_inspection_parameter_code on public.inspection_parameters;
create trigger set_inspection_parameter_code
before insert on public.inspection_parameters
for each row execute function public.trg_set_inspection_parameter_code();

update public.inspection_parameters
set parameter_code = 'PPM-' || upper(substr(id::text,1,8))
where category = 'PM Check'
  and nullif(trim(coalesce(parameter_code,'')),'') is null;

insert into public.inspection_parameters
  (category, parameter_name, description, severity, status)
select
  'PM Check',
  'Apakah PM check telah dilakukan pada tanggal xxx (sesuai yang di planning)',
  'Pilih Aman jika PM check sudah dilakukan sesuai tanggal planning. Pilih Tidak Aman jika belum dilakukan / tidak sesuai jadwal.',
  'High',
  'Aktif'
where not exists (
  select 1
  from public.inspection_parameters
  where category = 'PM Check'
    and parameter_name = 'Apakah PM check telah dilakukan pada tanggal xxx (sesuai yang di planning)'
);

notify pgrst, 'reload schema';

commit;
