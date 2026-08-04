-- Ejecutar esto en Supabase: Dashboard > SQL Editor > New query > Run

create table if not exists app_storage (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz not null default now()
);

-- Habilita Row Level Security
alter table app_storage enable row level security;

-- Política simple: cualquiera con la anon key puede leer y escribir.
-- Es lo mismo nivel de "protección" que tenía el artefacto original
-- (cualquiera con el link podía ver y modificar los datos).
-- Si más adelante querés restringirlo, reemplazá esto por políticas
-- basadas en auth.uid() y activá Supabase Auth.
create policy "allow all read" on app_storage
  for select using (true);

create policy "allow all write" on app_storage
  for insert with check (true);

create policy "allow all update" on app_storage
  for update using (true);
