"""
Script de migracion de una sola corrida: productos.xlsx + imagenes/ -> Supabase.

Uso:
    SUPABASE_URL=https://xxxx.supabase.co \
    SUPABASE_SERVICE_ROLE_KEY=xxxx \
    ADMIN_EMAIL=admin@ejemplo.com \
    ADMIN_PASSWORD=una-contrasena-larga \
    python migrate_to_supabase.py

Requiere que supabase/schema.sql ya haya sido ejecutado en el proyecto de Supabase.
La SERVICE_ROLE_KEY nunca debe subirse a git ni quedar en Netlify: solo se usa
en esta corrida local.
"""

import mimetypes
import os
import sys

import requests

from catalogo import leer_productos_excel

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
ADMIN_EMAIL = os.environ.get("ADMIN_EMAIL", "")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")

if not all([SUPABASE_URL, SERVICE_KEY]):
    sys.exit("Faltan SUPABASE_URL y/o SUPABASE_SERVICE_ROLE_KEY en las variables de entorno.")

HEADERS = {
    "apikey": SERVICE_KEY,
    "Authorization": f"Bearer {SERVICE_KEY}",
}


def subir_imagen(imagen_local, marca, codigo):
    """Sube la foto local a Storage y devuelve el path guardado, o '' si no hay foto real."""
    if imagen_local.endswith("logo-mucho-de-ti.jpg"):
        return ""

    extension = os.path.splitext(imagen_local)[1] or ".jpg"
    path = f"{marca}/{codigo}{extension}"
    tipo = mimetypes.guess_type(imagen_local)[0] or "image/jpeg"

    with open(imagen_local, "rb") as f:
        respuesta = requests.post(
            f"{SUPABASE_URL}/storage/v1/object/imagenes/{path}",
            headers={**HEADERS, "Content-Type": tipo, "x-upsert": "true"},
            data=f.read(),
            timeout=60,
        )

    if respuesta.status_code >= 300:
        print(f"  ! No se pudo subir {imagen_local}: {respuesta.status_code} {respuesta.text[:200]}")
        return ""

    return path


def upsert_producto(producto, imagen_path):
    fila = {
        "codigo": producto["codigo"],
        "nombre": producto["nombre"],
        "descripcion": producto["descripcion"],
        "rubro": producto["rubro"],
        "marca": producto["marca"],
        "precio": producto["precio"],
        "talles": producto["talles"],
        "stock": producto["stock"],
        "destacado": producto["destacado"],
        "oferta": producto["oferta"],
        "imagen_path": imagen_path,
    }

    respuesta = requests.post(
        f"{SUPABASE_URL}/rest/v1/productos",
        headers={
            **HEADERS,
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates",
        },
        json=fila,
        timeout=30,
    )

    if respuesta.status_code >= 300:
        print(f"  ! No se pudo guardar {producto['codigo']}: {respuesta.status_code} {respuesta.text[:200]}")
        return False

    return True


def crear_usuario_admin():
    if not ADMIN_EMAIL or not ADMIN_PASSWORD:
        print("ADMIN_EMAIL/ADMIN_PASSWORD no definidos: me salteo la creacion del usuario admin.")
        return

    respuesta = requests.post(
        f"{SUPABASE_URL}/auth/v1/admin/users",
        headers={**HEADERS, "Content-Type": "application/json"},
        json={"email": ADMIN_EMAIL, "password": ADMIN_PASSWORD, "email_confirm": True},
        timeout=30,
    )

    if respuesta.status_code >= 300:
        print(f"! No se pudo crear el usuario admin: {respuesta.status_code} {respuesta.text[:300]}")
    else:
        print(f"Usuario admin creado: {ADMIN_EMAIL}")


def main():
    productos = leer_productos_excel()
    print(f"Productos leidos de productos.xlsx: {len(productos)}")

    ok = 0
    for producto in productos:
        imagen_path = subir_imagen(producto["imagen"], producto["marca"], producto["codigo"])
        if upsert_producto(producto, imagen_path):
            ok += 1
        print(f"  {producto['codigo']} - {producto['nombre'][:40]}")

    print(f"\nProductos migrados con exito: {ok}/{len(productos)}")

    crear_usuario_admin()


if __name__ == "__main__":
    main()
