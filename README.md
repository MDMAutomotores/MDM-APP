# MDM Automotores — App de gestión de clientes

## Qué se arregló respecto a la versión anterior
1. **Exportar Excel**: ahora genera el archivo como `data:` URL en base64 en vez de blob (más confiable en navegadores móviles) y todo el proceso está envuelto en try/catch, así que si algo falla vas a ver un mensaje claro en vez de que no pase nada.
2. **Eliminar usuarios**: el panel de administración de usuarios no tenía implementado el borrado. Ahora cada usuario (que no sea admin ni vos mismo) tiene un botón de basurero con confirmación.
3. **Cámara y galería**: antes había un solo botón genérico de "agregar foto" que en muchos celulares no ofrecía la opción de cámara. Ahora hay dos botones separados — uno fuerza la cámara trasera (`capture="environment"`), el otro abre la galería.
4. **Persistencia real**: la app usaba `window.storage`, una API que solo existe dentro de Claude.ai. Para que funcione en Vercel con varios usuarios viendo los mismos datos, ahora usa **Vercel KV** (Redis) a través de una función serverless (`/api/storage.js`). El código de `App.jsx` no cambió su forma de guardar/leer datos — solo se reemplazó el "motor" por detrás en `src/main.jsx`.

## Logo
Puse un logo de relleno en `public/logo.jpg`. Reemplazalo por tu logo real (mismo nombre de archivo, o cambiá la constante `LOGO_MDM` en `src/App.jsx` si preferís otro nombre/formato).

## Pasos para subir a GitHub y desplegar en Vercel

### 1. Subir a GitHub
Desde la carpeta del proyecto:
```bash
git init
git add .
git commit -m "MDM Automotores - primera versión"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/NOMBRE_DEL_REPO.git
git push -u origin main
```
(Creá el repo vacío en GitHub primero, sin README ni .gitignore, para que el push no tenga conflictos.)

### 2. Importar en Vercel
1. Entrá a vercel.com → **Add New** → **Project**.
2. Elegí el repo que acabás de subir.
3. Vercel detecta automáticamente que es un proyecto Vite — no hace falta tocar la configuración de build.
4. Todavía **no** le des a Deploy. Primero conectá la base de datos (paso 3).

### 3. Conectar Vercel KV (la base de datos, gratis)
1. En el mismo proyecto en Vercel, andá a la pestaña **Storage**.
2. **Create Database** → elegí **KV** (o "Upstash for Redis", el nombre puede variar).
3. Seguí los pasos — Vercel conecta automáticamente las variables de entorno `KV_REST_API_URL` y `KV_REST_API_TOKEN` al proyecto, no hay que copiarlas a mano.

### 4. Deploy
Le das a **Deploy**. En un par de minutos tenés la URL pública (algo como `mdm-automotores.vercel.app`).

### 5. Primer uso
La primera vez que entren a la app, va a pedir configurar el usuario administrador (nombre + PIN de 4 dígitos). Desde ahí, ese admin puede crear el resto de los usuarios desde el ícono de personas (arriba a la derecha).

## Desarrollo local (opcional)
```bash
npm install
npm run dev
```
Nota: en local, `/api/storage` no va a funcionar a menos que uses `vercel dev` (que sí simula las funciones serverless) en vez de `npm run dev`, o que apuntes a variables de entorno de Vercel KV en un archivo `.env.local`.
