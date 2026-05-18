-- Food Index v47 - parameter metadata, dashboard filters, and safety hotfixes
-- Run this once in Supabase SQL Editor before testing v47.

alter table public.food_parameters
  add column if not exists standard_parameter text,
  add column if not exists hazard_code text,
  add column if not exists behavior text;

alter table public.food_findings
  add column if not exists due_date date;

-- Keep status constraints aligned with current Food Index flow.
alter table public.food_findings
  drop constraint if exists food_findings_status_check;

alter table public.food_findings
  add constraint food_findings_status_check
  check (
    status in (
      'Need Action Plan',
      'Action Plan Submitted',
      'Waiting Approval',
      'Approved',
      'Rejected',
      'Open',
      'Closed'
    )
  );

alter table public.food_weekly_tasks
  drop constraint if exists food_weekly_tasks_status_check;

alter table public.food_weekly_tasks
  add constraint food_weekly_tasks_status_check
  check (
    status in (
      'Open',
      'In Progress',
      'Need Action Plan',
      'Waiting Approval',
      'Approved',
      'Expired',
      'Rejected'
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
