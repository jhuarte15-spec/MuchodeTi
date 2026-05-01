# Mucho de Ti - Página mantenida desde Excel

## Qué cambió

- Ya no hace falta editar `productos.json`.
- Los productos se editan en `productos.xlsx`.
- La página muestra solo productos destacados.
- Las imágenes se vinculan automáticamente por código y marca.
- La portada tiene una imagen a la derecha sobre el mismo fondo rosado.

## Cómo cargar productos destacados

Abrí `productos.xlsx` y en la columna `Destacado` escribí:

```text
Sí
```

en los productos que quieras mostrar en la web.

Los productos con `No` quedan guardados en el Excel, pero no se muestran si `MOSTRAR_SOLO_DESTACADOS = True`.

## Cómo vincular imágenes

Guardá las fotos así:

```text
imagenes/
  Ana Grant/
    109. Ana Grant.jpg
  Aretha/
    123. Aretha.jpg
  BLUO/
    456. BLUO.jpg
  Deville/
  So Pink/
  Stylo/
```

La imagen debe empezar con el código del producto.

Ejemplo:

```text
109. Ana Grant.jpg
```

se vincula con el producto de código `109`.

## Cómo regenerar la página

Instalá openpyxl una sola vez:

```bash
pip install openpyxl
```

Después ejecutá:

```bash
python generar_pagina.py
```

Eso actualiza `index.html`.

## Imagen de portada

La imagen de la portada está en:

```text
imagenes/hero-producto.jpg
```

Podés reemplazarla por otra foto, manteniendo exactamente ese nombre.

## WhatsApp

Cambiá el número en `generar_pagina.py`:

```python
WHATSAPP_NUMERO = "5491112345678"
```
"# MuchodeTi"  
