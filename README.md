# Control de Stock — Despliegue en Vercel + GitHub + Supabase

Este proyecto ya está listo para publicar. El código del dashboard es el mismo
que tenías en el artefacto, solo que el guardado de datos (`window.storage`,
que únicamente existe dentro de Claude.ai) fue reemplazado por Supabase, que
sí funciona en un sitio público.

## 1. Crear el proyecto en Supabase

1. Entrá a https://supabase.com y creá una cuenta / iniciá sesión.
2. **New project** → elegí un nombre (ej. `control-de-stock`), una contraseña
   de base de datos (guardala, no se vuelve a mostrar) y una región cercana
   (ej. South America - São Paulo).
3. Esperá 1-2 minutos a que se aprovisione.
4. En el menú izquierdo: **SQL Editor** → **New query** → pegá el contenido
   del archivo `supabase_setup.sql` incluido acá → **Run**. Esto crea la
   tabla `app_storage` donde se va a guardar toda la info del dashboard.
5. En el menú izquierdo: **Project Settings** → **API**. Copiá:
   - **Project URL** → esto es `VITE_SUPABASE_URL`
   - **anon public key** → esto es `VITE_SUPABASE_ANON_KEY`

## 2. Subir el código a GitHub

1. Creá un repositorio nuevo en GitHub (puede ser privado o público), sin
   README ni .gitignore (ya vienen incluidos acá).
2. Desde esta carpeta, en tu terminal:
   ```bash
   git init
   git add .
   git commit -m "Control de stock - primera versión"
   git branch -M main
   git remote add origin https://github.com/TU-USUARIO/TU-REPO.git
   git push -u origin main
   ```

## 3. Probarlo en local (opcional pero recomendado)

```bash
npm install
cp .env.example .env
# Editá .env y pegá tu VITE_SUPABASE_URL y VITE_SUPABASE_ANON_KEY
npm run dev
```

Abrí la URL que te muestra la terminal (normalmente http://localhost:5173) y
confirmá que carga bien y que al subir un Excel los datos se guardan
(refrescá la página y deberían seguir ahí — eso confirma que Supabase está
funcionando).

## 4. Publicar en Vercel

1. Entrá a https://vercel.com e iniciá sesión con tu cuenta de GitHub.
2. **Add New** → **Project** → elegí el repositorio que acabás de subir.
3. Vercel va a detectar automáticamente que es un proyecto Vite (Framework
   Preset: Vite). No hace falta tocar el build command ni el output
   directory.
4. Antes de darle a **Deploy**, abrí **Environment Variables** y agregá:
   - `VITE_SUPABASE_URL` = (tu Project URL de Supabase)
   - `VITE_SUPABASE_ANON_KEY` = (tu anon public key de Supabase)
5. **Deploy**. En 1-2 minutos te da una URL tipo
   `https://control-de-stock-tuusuario.vercel.app`, pública, para compartir
   con cualquiera.

## 5. Actualizaciones futuras

Cada vez que quieras cambiar algo del dashboard:
1. Editás `src/App.jsx`.
2. `git add . && git commit -m "cambio tal" && git push`.
3. Vercel redeploya solo, automáticamente, en cuanto detecta el push.

## Notas importantes

- **Datos compartidos**: con esta configuración, todos los que entren al
  link ven y modifican los mismos datos (una sola "planilla" compartida),
  igual que pasaba dentro del artefacto de Claude. Si en algún momento
  querés que cada usuario tenga sus propios datos, hay que sumar Supabase
  Auth (login) y ajustar las políticas de RLS — avisame si llegás a
  necesitarlo y lo armamos.
- **anon key pública**: la `anon key` de Supabase está pensada para ir en el
  frontend (no es secreta como la `service_role key`, que nunca debe usarse
  acá). Aun así, como cualquiera puede escribir en la tabla, alguien con
  conocimientos técnicos podría borrar o alterar los datos manualmente. Para
  un uso interno de equipo esto normalmente es aceptable; si va a ser
  público y sensible, conviene sumar autenticación.
- **Dominio propio**: si más adelante querés algo como
  `stock.tuempresa.com` en vez del `.vercel.app`, se configura en Vercel →
  Project → Settings → Domains.
