# OutlookFocusRestorerWithUI

Este es un complemento VSTO para Microsoft Outlook (versión clásica de escritorio) que permite restaurar el foco en el último correo electrónico seleccionado, incluso después de cambiar de carpeta o realizar nuevas búsquedas.

## ✨ Características

- Guarda automáticamente el último correo seleccionado.
- Agrega un botón en la cinta de opciones de Outlook llamado **"Restaurar último correo"**.
- Al hacer clic en el botón, se vuelve a abrir el último correo en una nueva ventana.

## 🛠 Requisitos

- Visual Studio 2019 o 2022 con la carga de trabajo **Desarrollo de Office/SharePoint**.
- Microsoft Outlook instalado (versión Win32 clásica).
- .NET Framework 4.7 o superior.

## 🚀 Instalación

1. Clona este repositorio o descarga el ZIP y descomprímelo.
2. Abre el archivo `.sln` en Visual Studio.
3. Compila el proyecto (`Ctrl + Shift + B`).
4. Ejecuta el complemento con `F5` (esto abrirá Outlook con el complemento cargado).

## 🧪 Uso

1. Selecciona un correo en Outlook.
2. Cambia de carpeta o realiza otra búsqueda.
3. Haz clic en el botón **"Restaurar último correo"** en la cinta de opciones.
4. El correo previamente seleccionado se abrirá en una nueva ventana.

## 📦 Estructura del proyecto

- `ThisAddIn.cs`: lógica principal del complemento.
- `FocusRibbon.cs`: lógica del botón de la cinta.
- `FocusRibbon.Designer.cs`: diseño de la interfaz de la cinta.
- `FocusRibbon.xml`: definición de la cinta de opciones.

## 🤝 Contribuciones

¡Las contribuciones son bienvenidas! Puedes:

- Reportar errores o sugerencias en la sección de Issues.
- Hacer un fork y enviar un Pull Request con mejoras.

## 📄 Licencia

Este proyecto está bajo la licencia MIT. Puedes usarlo, modificarlo y distribuirlo libremente.

---

Desarrollado con ❤️ para mejorar la productividad en Outlook clásico.
