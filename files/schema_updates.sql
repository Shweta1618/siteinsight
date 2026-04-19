-- ============================================================
-- SiteInsight · Schema Updates
-- Run these in Supabase SQL Editor AFTER the original schema
-- ============================================================

-- 1. Add responsible_person to wpr_activities
alter table wpr_activities
  add column if not exists responsible_person text;

-- 2. Add status to detection_results
alter table detection_results
  add column if not exists status text
    check (status in ('Open','In Progress','Resolved'))
    default 'Open';

-- 3. New site_photos table
create table if not exists site_photos (
    id               serial primary key,
    week_number      int        not null,
    phase            text       not null,
    image_url        text       not null,
    caption          text,
    activity_ref     text,
    is_placeholder   boolean    default true,
    created_at       timestamptz not null default now()
);

create index if not exists idx_photos_week  on site_photos(week_number);
create index if not exists idx_photos_phase on site_photos(phase);

-- 4. Supabase Storage bucket for real site photos
-- Run this separately if you want real photo uploads later:
-- insert into storage.buckets (id, name, public)
-- values ('site-photos', 'site-photos', true)
-- on conflict do nothing;
