# REGALO PARTY — versión GitHub Pages

**Idea:** la parte bonita (HTML) se sirve desde GitHub Pages, así no aparece el
aviso de Google *"esta aplicación no ha sido verificada / creada por un usuario"*.
Apps Script se queda detrás como motor (API + Google Sheets). Los datos no se
mueven: siguen en la misma hoja de cálculo.

## Archivos

| Archivo | Dónde va |
|---|---|
| `index.html` | GitHub — página de los **participantes** |
| `admin.html` | GitHub — **panel de control** (pide PIN) |
| `estilos.css` | GitHub |
| `api.js` | GitHub — puente con Apps Script |
| `Codigo.gs` | **Apps Script** — sustituye al `Codigo.gs` actual |

`Estilos.html`, `Index.html` y `Admin.html` en Apps Script ya no se usan: puedes
dejarlos ahí sin tocar, no molestan.

## Pasos

### 1) Apps Script
1. Abre el proyecto y sustituye **todo** el contenido de `Codigo.gs` por el nuevo.
2. En la línea `const WEB_BASE = '...'` pon la URL real de tu GitHub Pages,
   **con la barra final**. Ejemplo: `https://puntofibra.github.io/regalo/`
3. Guarda → **Implementar › Gestionar implementaciones › ✏️ › Nueva versión › Implementar.**
   - *Ejecutar como:* **Yo**
   - *Quién tiene acceso:* **Cualquier usuario** ← imprescindible, si no la web de
     GitHub no podrá hablar con el script.
4. La URL `/exec` **no cambia** si actualizas la implementación existente.
   Si por lo que sea cambia, cópiala en `api.js` (variable `API_URL`).

### 2) GitHub
1. Repo nuevo, **público**, por ejemplo `regalo`.
2. Sube los 4 archivos (`index.html`, `admin.html`, `estilos.css`, `api.js`).
3. **Settings › Pages › Source: Deploy from a branch › `main` / `/ (root)` › Save.**
4. En 1–2 minutos tendrás:
   - Participantes → `https://puntofibra.github.io/regalo/`
   - Panel de control → `https://puntofibra.github.io/regalo/admin.html`

### 3) Limpieza opcional
En Apps Script, ejecuta una vez la función `limpiarEnlacesAntiguos()` para borrar
de la hoja CONFIG los enlaces viejos que causaban el bug. No es obligatorio.

## Qué se ha arreglado

**1. Los enlaces del producto que no se podían borrar.**
El panel guardaba la lista nueva en la clave `regaloEnlaces`, pero las claves
antiguas (`regaloEnlace`, `regaloEnlace2`, `regaloEnlace3`, `regaloEnlace4`)
seguían con los valores viejos en la hoja. Al leer, si la lista nueva quedaba
vacía el código hacía un "plan B" y volvía a coger esas claves antiguas → los
enlaces borrados reaparecían. Ahora:
- la lista nueva manda siempre, aunque esté vacía;
- al guardar se vacían las claves antiguas;
- el panel ya no rellena los enlaces con valores de respaldo.

**2. El texto de la descripción del regalo.**
Tenía `max-width: 520px` dentro de una tarjeta de 760px, así que quedaba
estrecho. Ahora ocupa todo el ancho y va justificado a ambos lados.

## Aviso sobre el PIN
El PIN viaja en la llamada al script (igual que antes). Sirve para que nadie
toque el panel por casualidad, pero no es seguridad de banco: no publiques la
URL de `admin.html` en sitios abiertos.
