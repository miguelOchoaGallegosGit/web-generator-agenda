# Generador de Agendas Word

Aplicacion web en React para generar agendas imprimibles en formato `.docx` a partir de un rango de fechas y uno de tres modelos visuales.

## Plan de trabajo

1. Base tecnica: React + Vite + TypeScript para una app rapida, simple de mantener y compatible con Vercel.
2. Experiencia de usuario: formulario compacto, miniaturas reales de los modelos, validacion inmediata y bloqueo de controles durante la generacion.
3. Reglas de negocio: la fecha fin no puede ser anterior a la fecha inicio y el rango maximo permitido es de un año.
4. Generacion Word: uso de `docx` en el navegador para crear el archivo sin backend.
5. Descarga: al terminar, el navegador descarga automaticamente el `.docx` y muestra un mensaje de agradecimiento.
6. Publicacion: repositorio en GitHub conectado a Vercel con build automatico.

## Modelos incluidos

- Clasica sobria: dos dias por pagina, lineas amplias y estilo limpio.
- Color pop: cuatro dias por pagina con bloques de color.
- Semana vista: una semana por pagina con acentos de color y mini calendario mensual.

## Comandos

```bash
npm install
npm run dev
npm run build
```

## Despliegue en Vercel

1. Subir el proyecto a un repositorio personal de GitHub.
2. En Vercel, crear un nuevo proyecto importando ese repositorio.
3. Verificar la configuracion:
   - Framework Preset: `Vite`
   - Build Command: `npm run build`
   - Output Directory: `dist`
4. Publicar. Vercel ejecutara el build y servira la app como sitio estatico.
