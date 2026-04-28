-- SQL PATCH FINAL V4
-- Tujuan:
-- 1) rapikan data applications agar tidak dobel di dropdown aplikasi;
-- 2) rename nama aplikasi ketiga menjadi "Inspeksi";
-- 3) pindahkan mapping user_app_access dari aplikasi duplikat ke aplikasi canonical;
-- 4) pastikan Platform Admin/App Admin yang sudah punya Sidak juga punya akses DRD Driver dan Inspeksi.

begin;

-- Canonicalize duplicated application rows.
do $$
declare
  v_sid uuid;
  v_drd uuid;
  v_ins uuid;
  r record;
begin
  select id into v_sid
  from public.applications
  where lower(coalesce(app_code,'')) like '%sidak%'
     or lower(coalesce(app_code,'')) like '%fatigue%'
     or lower(coalesce(app_name,'')) like '%sidak%'
     or lower(coalesce(app_name,'')) like '%fatigue%'
  order by created_at nulls last, id::text
  limit 1;

  if v_sid is null then
    insert into public.applications (app_code, app_name, description, status)
    values ('sidak_fatigue', 'Sidak Fatigue', 'Aplikasi sidak fatigue driver', 'Aktif')
    returning id into v_sid;
  end if;

  update public.applications
  set app_code='sidak_fatigue', app_name='Sidak Fatigue', description='Aplikasi sidak fatigue driver', status='Aktif'
  where id=v_sid;

  select id into v_drd
  from public.applications
  where lower(coalesce(app_code,'')) like '%drd%'
     or lower(coalesce(app_name,'')) like '%drd%'
  order by created_at nulls last, id::text
  limit 1;

  if v_drd is null then
    insert into public.applications (app_code, app_name, description, status)
    values ('drd_driver', 'DRD Driver', 'Aplikasi DRD dan test driver', 'Aktif')
    returning id into v_drd;
  end if;

  update public.applications
  set app_code='drd_driver', app_name='DRD Driver', description='Aplikasi DRD dan test driver', status='Aktif'
  where id=v_drd;

  select id into v_ins
  from public.applications
  where lower(coalesce(app_code,'')) like '%inspeksi%'
     or lower(coalesce(app_name,'')) like '%inspeksi%'
  order by created_at nulls last, id::text
  limit 1;

  if v_ins is null then
    insert into public.applications (app_code, app_name, description, status)
    values ('inspeksi_unit', 'Inspeksi', 'Aplikasi inspeksi unit dan kelayakan parkiran', 'Aktif')
    returning id into v_ins;
  end if;

  update public.applications
  set app_code='inspeksi_unit', app_name='Inspeksi', description='Aplikasi inspeksi unit dan kelayakan parkiran', status='Aktif'
  where id=v_ins;

  -- Pindahkan mapping dari aplikasi duplikat ke canonical Sidak.
  for r in
    select id from public.applications
    where id <> v_sid and (
      lower(coalesce(app_code,'')) like '%sidak%'
      or lower(coalesce(app_code,'')) like '%fatigue%'
      or lower(coalesce(app_name,'')) like '%sidak%'
      or lower(coalesce(app_name,'')) like '%fatigue%'
    )
  loop
    update public.user_app_access set app_id = v_sid where app_id = r.id;
    delete from public.applications where id = r.id;
  end loop;

  -- Pindahkan mapping dari aplikasi duplikat ke canonical DRD.
  for r in
    select id from public.applications
    where id <> v_drd and (
      lower(coalesce(app_code,'')) like '%drd%'
      or lower(coalesce(app_name,'')) like '%drd%'
    )
  loop
    update public.user_app_access set app_id = v_drd where app_id = r.id;
    delete from public.applications where id = r.id;
  end loop;

  -- Pindahkan mapping dari aplikasi duplikat ke canonical Inspeksi.
  for r in
    select id from public.applications
    where id <> v_ins and (
      lower(coalesce(app_code,'')) like '%inspeksi%'
      or lower(coalesce(app_name,'')) like '%inspeksi%'
    )
  loop
    update public.user_app_access set app_id = v_ins where app_id = r.id;
    delete from public.applications where id = r.id;
  end loop;
end $$;

-- Hapus mapping yang benar-benar duplikat setelah app_id disatukan.
delete from public.user_app_access a
using public.user_app_access b
where a.id > b.id
  and a.user_id = b.user_id
  and a.app_id = b.app_id
  and a.role = b.role
  and coalesce(a.site_id::text, '') = coalesce(b.site_id::text, '')
  and a.status = b.status;

-- Tambahkan akses DRD + Inspeksi untuk admin yang sudah punya akses Sidak.
with canonical_apps as (
  select app_code, id from public.applications where app_code in ('sidak_fatigue','drd_driver','inspeksi_unit')
), admin_sources as (
  select distinct u.user_id, u.role, u.site_id
  from public.user_app_access u
  join public.applications a on a.id = u.app_id
  where a.app_code = 'sidak_fatigue'
    and u.role in ('Platform Admin','App Admin')
    and u.status = 'Aktif'
), target_apps as (
  select id from canonical_apps where app_code in ('drd_driver','inspeksi_unit')
)
insert into public.user_app_access (user_id, app_id, role, site_id, status)
select s.user_id, t.id, s.role, s.site_id, 'Aktif'::public.user_status
from admin_sources s
cross join target_apps t
where not exists (
  select 1 from public.user_app_access x
  where x.user_id = s.user_id
    and x.app_id = t.id
    and x.role = s.role
    and coalesce(x.site_id::text,'') = coalesce(s.site_id::text,'')
);

-- Setelah data dobel beres, kunci app_code agar tidak dobel lagi.
create unique index if not exists uq_applications_app_code on public.applications(app_code);

notify pgrst, 'reload schema';

commit;
