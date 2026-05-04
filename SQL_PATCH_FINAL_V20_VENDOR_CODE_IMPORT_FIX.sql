-- =============================================================
-- PATCH FINAL V20 - VENDOR CODE IMPORT FIX
-- Tujuan: memastikan insert vendor dari UI/import tetap aman walaupun vendor_code wajib NOT NULL.
-- Jalankan di Supabase SQL Editor sebelum/bersamaan dengan deploy frontend pengganti.
-- =============================================================

create extension if not exists pgcrypto;

alter table if exists public.vendors
  add column if not exists vendor_code text;

create or replace function public.trg_set_vendor_code()
returns trigger
language plpgsql
as $$
begin
  if new.id is null then
    new.id := gen_random_uuid();
  end if;

  if nullif(trim(coalesce(new.vendor_code, '')), '') is null then
    new.vendor_code := 'VEN-' || upper(substr(new.id::text, 1, 8));
  else
    new.vendor_code := upper(trim(new.vendor_code));
  end if;

  return new;
end;
$$;

drop trigger if exists set_vendor_code on public.vendors;
create trigger set_vendor_code
before insert on public.vendors
for each row
execute function public.trg_set_vendor_code();

update public.vendors
set vendor_code = 'VEN-' || upper(substr(id::text, 1, 8))
where nullif(trim(coalesce(vendor_code, '')), '') is null;

create unique index if not exists vendors_vendor_code_key
on public.vendors (vendor_code);

alter table if exists public.vendors
  alter column vendor_code set not null;
