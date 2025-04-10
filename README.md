# 🖼️ Word Image Organizer (.docm)

This repository contains a Word Macro-Enabled Document (`.docm`) that automates the insertion and layout of multiple images from a folder directly into a Word document — perfect for quickly preparing print-ready layouts.

🌐 [English](#readme) | [Español](https://github.com/manuelnajera/image-organizer-word/blob/main/README-es.md#readme)

## 📂 What's Included

- `image-organizer.docm` — English version
- `organizador-imagenes.docm` — Spanish version

## ⚙️ How It Works

This tool contains a macro named `OrganizeImagesAndDelete` that:

1. Automatically detects and inserts all images located in the same folder as the `.docm` file.
2. Prompts you to set:
   - Number of **images per page** (1–20)
   - Number of **images per row** (only if more than 2 images per page)
3. Generates a new document called `OrganizedImages.docx` containing all the formatted images.
4. **Deletes the original images** from the folder to avoid clutter or duplicate usage.

⚠️ **Note:** The images are added to the new file immediately. If you don’t rename or move `OrganizedImages.docx`, it will be **overwritten the next time** the tool runs.

## 📦 Downloads

- [Download English Version](https://github.com/manuelnajera/image-organizer-word/releases/tag/v1.0.0-en)
- [Descargar Versión en Español](https://github.com/manuelnajera/image-organizer-word/releases/tag/v1.0.0-es)

## 🖥️ Platform Support

- ✅ Optimized for **Microsoft Word on Windows**
- ⚠️ On macOS, the button does not work. You must run the macro manually from **Developer → Macros → `OrganizeImagesAndDelete`**

## 🧩 Enabling Macros

You **must enable macros** when prompted in order to use the tool.

## 🔒 Disclaimer

Make sure you **want to delete** the images in the folder before running the macro. The deletion is automatic and cannot be undone.

## 📷 Example Use Case

> A client sends 15 images through WhatsApp. The cybercafé owner downloads them to a folder and runs this tool. In seconds, a neatly organized Word document is ready for printing, and the folder is clean.

## 🔐 License

This project is licensed under the [MIT License](LICENSE).  
You’re free to use, modify, or share it — but at your own risk.