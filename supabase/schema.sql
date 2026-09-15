-- Esquema de Supabase para MuchodeTi.
-- Se puede correr las veces que haga falta en el SQL Editor del proyecto de Supabase
-- (todas las sentencias son idempotentes: no falla si ya existen).
-- Reemplaza a productos.xlsx como fuente de verdad del catalogo.

create table if not exists productos (
  codigo text primary key,
  nombre text not null default '',
  descripcion text not null default '',
  rubro text not null default '',
  marca text not null default '',
  precio text not null default '',
  talles text not null default 'Consultar disponibilidad',
  stock text not null default '',
  destacado boolean not null default false,
  oferta text not null default '',
  imagen_path text not null default '',
  oculto boolean not null default false,
  updated_at timestamptz not null default now()
);

alter table productos add column if not exists oculto boolean not null default false;

alter table productos enable row level security;

drop policy if exists "Lectura publica de productos" on productos;
create policy "Lectura publica de productos"
  on productos for select
  to anon, authenticated
  using (true);

drop policy if exists "Solo el admin autenticado escribe productos" on productos;
create policy "Solo el admin autenticado escribe productos"
  on productos for all
  to authenticated
  using (true)
  with check (true);

-- Bucket publico para las fotos de producto.
insert into storage.buckets (id, name, public)
values ('imagenes', 'imagenes', true)
on conflict (id) do nothing;

drop policy if exists "Lectura publica de fotos" on storage.objects;
create policy "Lectura publica de fotos"
  on storage.objects for select
  to anon, authenticated
  using (bucket_id = 'imagenes');

drop policy if exists "Solo el admin autenticado sube/borra fotos" on storage.objects;
create policy "Solo el admin autenticado sube/borra fotos"
  on storage.objects for all
  to authenticated
  using (bucket_id = 'imagenes')
  with check (bucket_id = 'imagenes');
