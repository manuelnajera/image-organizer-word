# 🖼️ Organizador de Imágenes para Word (.docm)

Este repositorio contiene un documento habilitado para macros de Word (`.docm`) que automatiza la inserción y distribución de múltiples imágenes desde una carpeta, ideal para preparar archivos rápidamente para impresión.

🌐 [English](https://github.com/manuelnajera/image-organizer-word/blob/main/README.md#readme) | [Español](#readme)

## 📂 Contenido

- `organizador-imagenes.docm` — Versión en español
- `image-organizer.docm` — Versión en inglés

## ⚙️ ¿Cómo Funciona?

El archivo contiene una macro llamada `OrganizeImagesAndDelete` que:

1. Detecta automáticamente todas las imágenes ubicadas en la misma carpeta que el archivo `.docm`.
2. Solicita los siguientes parámetros:
   - Número de **imágenes por página** (entre 1 y 20)
   - Número de **imágenes por fila** (solo si son más de 2 por página)
3. Crea un documento nuevo llamado `ImagenesOrganizadas.docx` con las imágenes ya distribuidas.
4. **Elimina las imágenes originales** de la carpeta para evitar acumulaciones o duplicaciones.

⚠️ **Nota:** Las imágenes se insertan inmediatamente en el nuevo documento. Si no lo renombras o lo mueves, será **sobrescrito la próxima vez** que uses la herramienta.

## 🖥️ Compatibilidad

- ✅ Optimizado para **Microsoft Word en Windows**
- ⚠️ En macOS, el botón no funciona. Debes ejecutar la macro manualmente desde **Desarrollador → Macros → `OrganizeImagesAndDelete`**

## 🧩 Habilitar Macros

Al abrir el documento, debes **aceptar habilitar macros** para que la herramienta funcione correctamente.

## 🔐 Licencia

Este proyecto está licenciado bajo la [Licencia MIT](LICENSE).  
Puedes usarlo, modificarlo y compartirlo libremente, pero bajo tu propio riesgo.