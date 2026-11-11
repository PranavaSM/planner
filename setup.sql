-- Run once in Supabase: SQL Editor → New query → paste → Run.
create table if not exists public.items (
  id uuid primary key,
  user_id uuid not null default auth.uid() references auth.users (id) on delete cascade,
  kind text not null,
  title text not null,
  date text not null default '',
  time text not null default '',
  done boolean not null default false,
  created bigint not null
);

alter table public.items enable row level security;

drop policy if exists "Only you can see and change your items" on public.items;
create policy "Only you can see and change your items" on public.items
  for all to authenticated
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);
