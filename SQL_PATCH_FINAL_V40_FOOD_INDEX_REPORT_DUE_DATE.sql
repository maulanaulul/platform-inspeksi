-- Food Index v40 - Report temuan + Due Date follow up GL
-- Jalankan sekali di Supabase SQL Editor sebelum/bersamaan replace FoodIndexApp.jsx.

alter table if exists public.food_findings
  add column if not exists due_date date;

comment on column public.food_findings.due_date is 'Due date tindak lanjut temuan Food Index yang diisi GL bersama corrective dan preventive action.';

create index if not exists idx_food_findings_due_date
  on public.food_findings(due_date);

-- Pastikan status flow Food Index tetap mendukung action plan dan close outstanding.
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
