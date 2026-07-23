# Deploy — Vercel + Supabase

El panel corre como sitio estático (`index.html`) + funciones serverless de Python
(`api/*.py`) en Vercel. Como las funciones serverless son **efímeras** y **no
comparten disco**, el estado (archivos subidos, reporte generado, versión) se guarda
en **Supabase Storage**, no en `/tmp`.

## 1. Configurar Supabase

1. En tu proyecto de Supabase → **Storage** → crea un bucket llamado `ocupacion`
   (puede ser privado; las funciones usan la *service key*).
2. Ve a **Settings ▸ API** y copia:
   - **Project URL** → ej. `https://abcd1234.supabase.co`
   - **service_role key** (secreta, no la anon key)

Objetos que el sistema crea dentro del bucket:

| Objeto                  | Contenido                          |
|-------------------------|------------------------------------|
| `uploads/base.xlsx`     | Stock físico subido                |
| `uploads/cap.xlsx`      | Tabla de capacidades (opcional)    |
| `report/reporte.html`   | Reporte generado (lo sirve el panel)|
| `report/reporte.xlsx`   | Reporte en Excel                   |
| `state.json`            | Versión, timestamp, etiqueta, nombres |

## 2. Variables de entorno en Vercel

En el proyecto de Vercel → **Settings ▸ Environment Variables**, agrega:

| Variable                | Valor                                  |
|-------------------------|----------------------------------------|
| `SUPABASE_URL`          | Project URL de Supabase                |
| `SUPABASE_SERVICE_KEY`  | service_role key                       |
| `SUPABASE_BUCKET`       | `ocupacion` (opcional, es el default)  |

Aplícalas a Production (y Preview si usas ramas). Luego **redeploy**.

## 3. Deploy

```bash
git add -A
git commit -m "deploy: persistencia en Supabase Storage"
git push
```

Vercel detecta `api/*.py` automáticamente (runtime Python + `requirements.txt`).
`vercel.json` fija los tiempos máximos de ejecución (el análisis usa `maxDuration: 60`
porque pandas/openpyxl tardan más que el default de 10 s).

## Notas

- El análisis se ejecuta **en el mismo proceso** (importando `analisis_ocupacion_sep.py`),
  no por `subprocess`. Por eso ese archivo se incluye en el bundle vía `includeFiles`.
- Si ves *"Faltan variables de entorno SUPABASE_URL / SUPABASE_SERVICE_KEY"* en el panel,
  faltan las env vars o no se hizo redeploy tras agregarlas.
- Límite de subida de Vercel: el body de una función serverless es ~4.5 MB. El Excel de
  ejemplo pesa ~3 MB, así que entra; si algún archivo fuera mayor, habría que subir
  directo a Supabase desde el navegador.
