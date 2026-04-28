-- SQL_PATCH_FINAL_V4B_DEDUP_ACCESS_FIXED.sql
-- Fix duplicate application/user_app_access mapping without updating into unique-key conflicts.
-- Run this instead of the previous V4 patch if you got uq_user_app_access_active duplicate key error.

begin;

-- 1) Normalize application display names.
update public.applications
set app_name = 'Sidak Fatigue'
where lower(coalesce(app_name, '')) like '%sidak%';

update public.applications
set app_name = 'DRD Driver'
where lower(coalesce(app_name, '')) like '%drd%';

update public.applications
set app_name = 'Inspeksi'
where lower(coalesce(app_name, '')) like '%inspeksi%'
  and lower(coalesce(app_name, '')) not like '%sidak%'
  and lower(coalesce(app_name, '')) not like '%drd%';

-- 2) Build canonical app map. If duplicate rows exist for the same app_name,
--    keep one canonical id and move access safely by insert-then-delete, not by update.
drop table if exists pg_temp.tmp_app_canonical;
create temp table tmp_app_canonical as
with canonical as (
  select
    app_name,
    (array_agg(id order by id::text))[1] as canonical_id
  from public.applications
  where app_name in ('Sidak Fatigue', 'DRD Driver', 'Inspeksi')
  group by app_name
)
select
  a.id as old_id,
  a.app_name,
  c.canonical_id
from public.applications a
join canonical c on c.app_name = a.app_name
where a.app_name in ('Sidak Fatigue', 'DRD Driver', 'Inspeksi');

-- 3) Insert missing canonical access rows first.
--    This prevents duplicate-key conflict when duplicate app ids are merged.
insert into public.user_app_access (user_id, app_id, role, site_id, status)
select distinct on (u.user_id, m.canonical_id, u.role, coalesce(u.site_id, '00000000-0000-0000-0000-000000000000'::uuid))
  u.user_id,
  m.canonical_id,
  u.role,
  u.site_id,
  coalesce(u.status, 'Aktif'::public.user_status) as status
from public.user_app_access u
join tmp_app_canonical m on m.old_id = u.app_id
where not exists (
  select 1
  from public.user_app_access x
  where x.user_id = u.user_id
    and x.app_id = m.canonical_id
    and x.role = u.role
    and coalesce(x.site_id, '00000000-0000-0000-0000-000000000000'::uuid)
        = coalesce(u.site_id, '00000000-0000-0000-0000-000000000000'::uuid)
);

-- 4) Delete access rows that point to non-canonical duplicate app ids.
delete from public.user_app_access u
using tmp_app_canonical m
where u.app_id = m.old_id
  and m.old_id <> m.canonical_id;

-- 5) Remove duplicate canonical access rows if any already exist.
delete from public.user_app_access u
using (
  select
    id,
    row_number() over (
      partition by user_id, app_id, role, coalesce(site_id, '00000000-0000-0000-0000-000000000000'::uuid)
      order by
        case when status = 'Aktif'::public.user_status then 0 else 1 end,
        id
    ) as rn
  from public.user_app_access
) d
where u.id = d.id
  and d.rn > 1;

-- 6) Deactivate duplicate app rows if applications.status exists, instead of deleting them.
--    This avoids foreign-key errors while keeping selectors clean through user_app_access.
do $$
begin
  if exists (
    select 1 from information_schema.columns
    where table_schema = 'public'
      and table_name = 'applications'
      and column_name = 'status'
  ) then
    execute $q$
      update public.applications a
      set status = 'Nonaktif'
      from tmp_app_canonical m
      where a.id = m.old_id
        and m.old_id <> m.canonical_id
    $q$;
  end if;
end $$;

notify pgrst, 'reload schema';

commit;
