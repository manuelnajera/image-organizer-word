# 🖼️ Organizador de Imágenes para Word (.docm)

[![GitHub issues](https://img.shields.io/github/issues/manuelnajera/image-organizer-word)](https://github.com/manuelnajera/image-organizer-word/issues)
[![GitHub license](https://img.shields.io/github/license/manuelnajera/image-organizer-word)](https://github.com/manuelnajera/image-organizer-word/blob/main/LICENSE)
[![GitHub Sponsors](https://shields.io/badge/Github%20Sponsors-Support%20me-blue?logo=GitHub+Sponsors)](https://github.com/sponsors/manuelnajera "Apóyame en GitHub Sponsors")

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

## 📦 Descargas
- [Descargar Versión en Español](https://github.com/manuelnajera/image-organizer-word/releases/tag/v1.0.0-es)
- [Download English Version](https://github.com/manuelnajera/image-organizer-word/releases/tag/v1.0.0-en)

## 🖥️ Compatibilidad

- ✅ Optimizado para **Microsoft Word en Windows**
- ⚠️ En macOS, el botón no funciona. Debes ejecutar la macro manualmente desde **Desarrollador → Macros → `OrganizeImagesAndDelete`**

## 🧩 Habilitar Macros

Al abrir el documento, debes **aceptar habilitar macros** para que la herramienta funcione correctamente.

## 🔒 Aviso importante

Asegurate de que **quieres eliminar** Las imágenes en la carpeta antes de ejecutar la macro. La eliminación es automática y no se puede deshacer.

## 📷 Ejemplo de uso

> Un cliente envía 15 imágenes a través de WhatsApp. El propietario del CyberCafé los descarga a una carpeta y ejecuta esta herramienta. En segundos, un documento de Word bien organizado está listo para la impresión, y la carpeta está limpia.

## 🔐 Licencia

Este proyecto está licenciado bajo la [Licencia MIT](LICENSE).  
Puedes usarlo, modificarlo y compartirlo libremente, pero bajo tu propio riesgo.