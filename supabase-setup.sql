-- Ejecutar en Supabase > SQL Editor

create table if not exists deposito_state (
  id          integer primary key default 1,
  articles    jsonb   not null default '[]',
  deliveries  jsonb   not null default '[]',
  movimientos jsonb   not null default '[]',
  next_id     integer not null default 1,
  next_del    integer not null default 1,
  next_mov    integer not null default 1,
  updated_at  timestamptz default now()
);

-- Insertar la fila única si no existe
insert into deposito_state (id) values (1)
on conflict (id) do nothing;

-- Deshabilitar RLS (app interna, no necesita restricciones por usuario)
alter table deposito_state disable row level security;
