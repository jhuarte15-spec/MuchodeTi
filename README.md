"# MuchodeTi"

Sitio: https://muchodeti.netlify.app/

## Cómo se administra el catálogo

La fuente de verdad del catálogo es una base de datos en **Supabase** (tabla `productos` + bucket de fotos `imagenes`), editable desde el panel privado que le dieron a la administradora (URL secreta `admin-<token>.html`, sin login: al abrirla queda conectada automáticamente).

Cada vez que se guarda un cambio en el panel, se dispara automáticamente un **Build Hook** de Netlify que vuelve a generar `index.html` (corriendo [generar_pagina.py](generar_pagina.py), que ahora lee de Supabase en vez del Excel) y lo publica. El cambio se ve en el sitio en aproximadamente 1 minuto.

`productos.xlsx` y la carpeta `imagenes/` **ya no alimentan el sitio**: quedan como respaldo histórico de cómo estaba cargado el catálogo antes de migrar a Supabase (ver [migrate_to_supabase.py](migrate_to_supabase.py), el script que hizo esa migración una única vez).

### Estructura del código

- `catalogo.py` — funciones compartidas: lectura de productos (desde Supabase, o desde el Excel para la migración histórica), clasificación por categoría, formato de precios.
- `generar_pagina.py` — genera `index.html` a partir de los productos de Supabase. Lo corre Netlify en cada build (ver `netlify.toml`).
- `migrate_to_supabase.py` — script de una sola corrida, ya ejecutado, que subió el Excel y las fotos a Supabase.
- `admin-<token>.html` — panel de administración. No se linkea desde el sitio público; el token de la URL es lo que lo mantiene privado.
- `supabase/schema.sql` — definición de la tabla y las políticas de seguridad (RLS) del proyecto de Supabase.
