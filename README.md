# ReCéTaMe

Recetario personal: raciones ajustables, macros por ración, modo cocina y recetas editables desde el móvil.

- `index.html`: la app (GitHub Pages).
- `servidor/Codigo.gs`: copia del servidor de Google Apps Script. El que funciona es el pegado en la hoja «ReCéTaMe datos»; este es solo una copia de seguridad.
- Las recetas viven en la hoja de Google y las fotos en la carpeta de Drive «ReCéTaMe fotos».
- `datos/nuevas.json`: recetas nuevas que sube Claude. El servidor las copia a la hoja automáticamente (cada receta lleva `id` y `revision`; cada revisión se aplica una sola vez).
- `img/recetas/`: fotos de esas recetas nuevas.
