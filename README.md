# Apps Script – Reservaciones a Google Sheets

## `src/OwnerRez.js`

Rastrea **todos** los correos de `oru5b14b666b9x@inquiryspot.com` (notificaciones de La Paz Bay Rentals / OwnerRez), extrae la reserva y la guarda en la hoja **RESERVAS** (configurable en `OWNERREZ_CONFIG.SHEET_NAME`). Las reservas nuevas se agregan **al final, en filas nuevas**; el script no cambia el formato de las filas que ya existen.

| Col | Campo | Origen en el correo |
|-----|-------|---------------------|
| A | Unidad | Nombre de la propiedad (línea antes de *Check-in*) |
| B | Guest | Guest Name |
| C | Check-in | fecha real (`dd/mm/yyyy`) |
| D | Check-out | fecha real (`dd/mm/yyyy`) |
| E | Number of nights | |
| F | Number of guests | |
| G | Confirmation code | llave de la reserva |
| H | Source | |
| I | status | `Confirmada` / `Modificado` (naranja) / `Cancelada` (rojo), según el asunto |
| J | fechaMail | fecha y hora del correo (se muestra `dd/mm/yyyy`) |

### Cómo funciona
- **Primera ejecución:** recorre todo el histórico del remitente. Si no termina en ~4.5 min, guarda en qué punto se quedó y sigue en la siguiente ejecución.
- **Después:** solo revisa los correos recientes (desde la última ejecución, con 2 días de margen).
- **Sin duplicados:** cada reserva se identifica por su *Confirmation code*. Un correo solo modifica la fila si es **más nuevo** que el último aplicado. Así, una cancelación o modificación actualiza la misma fila sin importar el orden en que se lean los correos.
- Si la hoja queda vacía o se cambia `SHEET_NAME`, la siguiente ejecución vuelve a recorrer todo el histórico.
- Si un correo nuevo trae un campo vacío, se conserva el valor anterior.
- Los correos que no son reservas (p. ej. *inquiries*) se omiten y quedan registrados en el log.

### Instalación
1. En la Google Sheet: *Extensiones → Apps Script*, crea un archivo `OwnerRez.gs` y pega el contenido de `src/OwnerRez.js`.
   Todas las funciones llevan el prefijo `ownerRez`, así que puede convivir con tus otros scripts en el mismo proyecto.
2. Agrega el menú en tu `onOpen` existente:
   ```js
   function onOpen() {
     // ...tu menú actual
     ownerRezAddMenu();
   }
   ```
   (Si el proyecto no tiene `onOpen`, créalo solo con esa línea.)
3. Ejecuta **`ownerRezDebugLatest`** una vez. Pide permisos y muestra en el log lo que extrae de los últimos 5 correos, sin escribir nada.
4. Ejecuta **`ownerRezSync`** para la carga inicial.
5. Ejecuta **`ownerRezInstallTrigger`** una vez para que se ejecute automáticamente cada hora.

`ownerRezResetBackfill` fuerza a recorrer de nuevo todo el histórico. No borra filas.

### Pruebas locales
```bash
npm test
```
Prueban el parser con el template real (`tests/fixtures/new-reservation.html`) y la sincronización con stubs de Gmail/Sheets.
