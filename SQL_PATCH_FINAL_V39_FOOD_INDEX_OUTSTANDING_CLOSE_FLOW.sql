-- SQL_PATCH_FINAL_V39_FOOD_INDEX_OUTSTANDING_CLOSE_FLOW.sql
-- Food Index v39
-- - Pastikan status lifecycle Food Index mendukung flow:
--   Inspeksi -> Need Action Plan -> Waiting Approval -> Approved
--   Outstanding -> Open -> Waiting Close Approval -> Closed / Rejected
-- - Pastikan function expire task mengenali status Need Action Plan

alter table public.food_weekly_tasks
  drop constraint if exists food_weekly_tasks_status_check;

alter table public.food_weekly_tasks
  add constraint food_weekly_tasks_status_check
  check (
    status in (
      'Open',
      'In Progress',
      'Need GL Action',
      'Need Action Plan',
      'Waiting Approval',
      'Approved',
      'Expired',
      'Rejected'
    )
  );

alter table public.food_findings
  drop constraint if exists food_findings_status_check;

alter table public.food_findings
  add constraint food_findings_status_check
  check (
    status in (
      'Open',
      'Need Action Plan',
      'Action Plan Submitted',
      'Validated',
      'Waiting Approval',
      'Approved',
      'Rejected',
      'Closed'
    )
  );

alter table public.food_outstandings
  drop constraint if exists food_outstandings_status_check;

alter table public.food_outstandings
  add constraint food_outstandings_status_check
  check (
    status in (
      'Open',
      'Action Plan Submitted',
      'Waiting Close Approval',
      'Closed',
      'Rejected'
    )
  );

alter table public.food_outstandings
  drop constraint if exists ck_food_outstanding_close_photos;

alter table public.food_outstandings
  add constraint ck_food_outstanding_close_photos
  check (
    status not in ('Waiting Close Approval','Closed')
    or (corrective_photo_url is not null and preventive_photo_url is not null)
  );

create or replace function public.expire_old_food_tasks(p_today date default current_date)
returns integer
language plpgsql
security definer
as $$
declare
  v_count integer;
begin
  update public.food_weekly_tasks
  set status = 'Expired', expired_at = now(), updated_at = now()
  where week_end_date < p_today
    and status in ('Open','In Progress','Need GL Action','Need Action Plan','Waiting Approval','Rejected');

  get diagnostics v_count = row_count;
  return v_count;
end;
$$;
