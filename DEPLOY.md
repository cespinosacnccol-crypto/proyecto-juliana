# Cómo desplegar en Streamlit Cloud (gratis)

## 1. Crear un repositorio en GitHub

1. Ve a https://github.com/new
2. Nombre del repo: `proyecto-juliana` (o el que quieras)
3. Clic en **Create repository**

## 2. Subir los archivos

Ejecuta estos comandos en PowerShell (en la carpeta del proyecto):

```powershell
cd "C:\Users\Catherine Espinosa\Desktop\PROYECTO JULIANA"
git init
git add -A
git commit -m "Primer commit: sistema de calificacion + web app"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/proyecto-juliana.git
git push -u origin main
```

*(Reemplaza `TU_USUARIO` por tu usuario de GitHub)*

## 3. Desplegar en Streamlit Cloud

1. Ve a https://streamlit.io/cloud
2. Inicia sesión con tu cuenta de GitHub
3. Clic en **New app**
4. Selecciona el repositorio `proyecto-juliana`
5. Branch: `main`
6. Main file path: `app.py`
7. Clic en **Deploy**

En ~2 minutos tu app estará en línea con un link como:
```
https://proyecto-juliana.streamlit.app
```

## 4. Compartir el link

Ese link es el que compartes. Los usuarios pueden:
- Subir archivos .xlsx
- Ver resultados en tabla
- Descargar los Excel calificados
- Descargar el acumulado general

## Notas

- Los archivos subidos **no se guardan en el servidor** (solo en memoria durante la sesión)
- Los resultados se pueden descargar inmediatamente
- Funciona en celulares, tablets y computadores
