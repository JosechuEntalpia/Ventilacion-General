# API Ventilación Genérica IEE

Este paquete mantiene los formatos originales:

- `plantilla.xlsx`: copia intacta del Excel subido.
- `plantilla.docx`: copia intacta del Word subido.
- `google-sites.html`: tabla HTML generada desde `Hoja1 (2)` manteniendo geometría, celdas combinadas, colores principales, alineación y estructura de la hoja.
- `server.js`: API Node/Express para Render.

## Render

Build command:

```bash
npm install
```

Start command:

```bash
npm start
```

## Google Sites

Pegar el contenido de `google-sites.html` en Insertar → Insertar código.

Cambiar esta línea:

```js
const API_BASE = "https://TU-SERVICIO-RENDER.onrender.com";
```

por la URL real de Render.

## Variables sin fórmula definida en Excel

Estas variables están en la hoja de variables, pero en la hoja de lógica no tienen fórmula asociada:

- MOlfPer11
- MOlfPer12
- MOlfPer21
- MOlfPer22

No se ha inventado fórmula para ellas. Se devuelven vacías mientras no se defina la lógica exacta.


## Corrección para Error API: Failed to fetch

El archivo `google-sites.html` incluye ahora un campo visible `URL API Render` y un botón `Probar API`.

Antes de calcular, pega ahí la URL real de Render, por ejemplo:

https://api-ventilacion-iee.onrender.com

Después pulsa `Probar API`. Debe mostrar `API conectada correctamente`.

Si no conecta, abre directamente en navegador:

https://TU-SERVICIO-RENDER.onrender.com/api/health

Debe devolver:

{"ok":true}


## Desplegables

DROPDOWNS APLICADOS

Se han respetado las validaciones reales existentes en el Excel original:

- B5 -> opciones J4:J6 bajo encabezado J3 = ODA
- B7 -> opciones L4:L5 bajo encabezado L3 = Edificio
- C11 -> opciones K4:K7 bajo encabezado K3 = IDA
- C17 -> opciones K4:K7 bajo encabezado K3 = IDA

Nota técnica: en el Excel recibido, la celda B6 contiene el valor 400 y NO tiene validación de lista; por eso no se ha convertido B6 en desplegable para no alterar el funcionamiento ni inventar una relación que no existe en el archivo.
