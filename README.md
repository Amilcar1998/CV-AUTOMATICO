# 📄 CV-AUTOMATICO: Generador Automatizado de Currículum Vitae Profesional

[![Python Version](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![python-docx](https://img.shields.io/badge/python--docx-v1.1.0%2B-blue?style=for-the-badge)](https://python-docx.readthedocs.io/)
[![Format](https://img.shields.io/badge/Output-DOCX-2B579A?style=for-the-badge&logo=microsoftword&logoColor=white)](https://www.microsoft.com/word)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)

**CV-AUTOMATICO** es una herramienta en Python diseñada para generar currículums vitae modernos, elegantes y profesionales en formato **Microsoft Word (.docx)** a partir de un archivo estructurado en **JSON**. 

Permite separar el **contenido** (datos personales, experiencia, habilidades) del **diseño y maquetación**, logrando un diseño de dos columnas asimétricas con colores corporativos, bordes personalizados, hipervínculos nativos y control estricto de saltos de página.

---

## 📑 Tabla de Contenidos

1. [Características Destacadas](#-características-destacadas)
2. [Estructura del Proyecto](#-estructura-del-proyecto)
3. [Estructura de Datos (cv_data.json)](#-estructura-de-datos-cv_datajson)
4. [Requisitos Previos](#-requisitos-previos)
5. [Instalación y Configuración](#-instalación-y-configuración)
6. [Uso y Generación](#-uso-y-generación)
7. [Personalización de Diseño y Estilos](#-personalización-de-diseño-y-estilos)
8. [Detalles Técnicos y XML Nativo](#-detalles-técnicos-y-xml-nativo)
9. [Contribución](#-contribución)
10. [Licencia](#-licencia)

---

## ✨ Características Destacadas

- 🎨 **Diseño Moderno Asimétrico**: Columna lateral izquierda (*Sidebar*) en gris platino para foto, datos de contacto, habilidades e idiomas, y columna principal derecha para perfil, trayectoria y educación.
- 🔷 **Bloque de Cabecera Contrastado**: Encabezado en azul profundo (*Teal Oscuro*) con tipografía blanca y divisores estilizados.
- 📊 **Desacoplamiento de Datos**: Toda la información se administra desde `cv_data.json` sin necesidad de modificar el código fuente en Python.
- 🖼️ **Foto de Perfil Integrada**: Detección e incrustación automática de `img.jpg` con fallback inteligente.
- 🔗 **Hipervínculos Nativos de Word**: Integración de links clickeables (LinkedIn, portafolios, correos) directamente en el documento generado.
- 📄 **Control Anti-Desborde y Paginación Limpia**:
  - Evita cortes bruscos en elementos de lista y bloques de experiencia (`keep_lines_together`, `keep_with_next`).
  - Habilita la división fluida de tablas entre páginas (`w:cantSplit = 0`).
  - Márgenes optimizados para hoja tamaño **A4**.

---

## 📂 Estructura del Proyecto

```text
CV-AUTOMATICO/
├── cv.py                 # Motor de renderizado en Python (python-docx + OpenXML)
├── cv_data.json          # Archivo de datos estructurado en formato JSON
├── img.jpg               # Fotografía de perfil para el currículum
├── CV_Eliseo_Lopez.docx  # Documento Word generado listo para enviar/imprimir
└── README.md             # Documentación completa del proyecto
```

---

## 🗂️ Estructura de Datos (cv_data.json)

El archivo `cv_data.json` organiza la información en secciones modulares:

```json
{
  "nombre_completo": "NOMBRE APELLIDO",
  "titulo_profesional": "Senior QA & Integration Specialist",
  "contacto": {
    "titulo": "📞 CONTACTO",
    "datos": {
      "telefono": "📱 +123 456 7890",
      "email": "✉️ usuario@email.com",
      "ubicacion": "📍 Ciudad, País",
      "linkedin": {
        "texto": "linkedin.com/in/usuario",
        "url": "https://www.linkedin.com/in/usuario"
      }
    }
  },
  "resumen_profesional": [
    "Párrafo 1 de resumen profesional...",
    "Párrafo 2 con especialidades y trayectoria..."
  ],
  "fortalezas": [
    "Liderazgo de proyectos",
    "Arquitectura de software",
    "Automatización de pruebas QA"
  ],
  "tecnologias": {
    "titulo": "💻 TECNOLOGÍAS",
    "lista": ["Python", "JavaScript", "Cypress", "SQL", "Git", "Docker"]
  },
  "competencias": {
    "titulo": "⚙️ COMPETENCIAS",
    "lista": ["Pensamiento crítico", "Resolución de problemas", "Trabajo en equipo"]
  },
  "idiomas": {
    "titulo": "🌐 IDIOMAS",
    "datos": [
      { "idioma": "Español", "nivel": "Nativo" },
      { "idioma": "Inglés", "nivel": "B2 / Intermedio-Avanzado" }
    ]
  },
  "otros_estudios": {
    "titulo": "📜 CERTIFICACIONES",
    "lista": ["Certificación QA Automation", "Scrum Master Professional"]
  },
  "formacion": {
    "titulo": "🎓 FORMACIÓN ACADÉMICA",
    "datos": [
      {
        "titulo": "Licenciatura / Ingeniería en Sistemas",
        "institucion": "Universidad Nacional",
        "ubicacion_periodo": "2018 - 2022",
        "nota": "Graduado con honores"
      }
    ]
  },
  "experiencia": {
    "titulo": "💼 EXPERIENCIA LABORAL",
    "datos": [
      {
        "empresa": "EMPRESA TECNOLÓGICA S.A.",
        "roles": [
          {
            "puesto": "Senior QA Automation Engineer",
            "periodo": "2022 - Presente",
            "descripcion": [
              "Diseño e implementación de frameworks de pruebas con Cypress y BrowserStack.",
              "Automatización de suites de regresión reduciendo los tiempos de release en un 40%."
            ],
            "tecnologias": "Cypress, JavaScript, BrowserStack, GitHub Actions"
          }
        ]
      }
    ]
  },
  "proyectos_destacados": {
    "titulo": "🚀 PROYECTOS DESTACADOS",
    "datos": [
      {
        "nombre": "Framework de Pruebas E2E Multi-Cloud",
        "descripcion": "Infraestructura de pruebas distribuida en BrowserStack y LambdaTest.",
        "tecnologias": "Cypress, Node.js, Mocha"
      }
    ]
  },
  "referencias_laborales": {
    "titulo": "👥 REFERENCIAS LABORALES",
    "datos": [
      {
        "nombre": "Lic. Juan Pérez",
        "puesto": "Gerente de Calidad",
        "empresa": "Tech Solutions",
        "telefono": "+123 456 7890",
        "correo": "juan.perez@email.com"
      }
    ]
  },
  "referencias_personales": {
    "titulo": "🤝 REFERENCIAS PERSONALES",
    "datos": [
      {
        "nombre": "Ing. María Gómez",
        "puesto": "Desarrolladora Senior",
        "telefono": "+123 456 7891",
        "correo": "maria.gomez@email.com"
      }
    ]
  }
}
```

---

## ⚙️ Requisitos Previos

- **Python**: Versión 3.8 o superior instalada.
- **Gestor de paquetes pip**.

---

## 📥 Instalación y Configuración

### 1. Clonar el repositorio
```bash
git clone https://github.com/Amilcar1998/CV-AUTOMATICO.git
cd CV-AUTOMATICO
```

### 2. Instalar las dependencias
```bash
pip install python-docx
```

---

## 🚀 Uso y Generación

1. **Actualiza tus datos**: Edita `cv_data.json` con tu información personal y trayectoria.
2. **Añade tu foto** (Opcional): Guarda tu foto con el nombre `img.jpg` en la raíz del proyecto.
3. **Ejecuta el script generador**:
   ```bash
   python cv.py
   ```
4. **Resultado**: Se creará el archivo Word correspondiente (ej. `CV_Eliseo_Lopez.docx`), completamente formateado y listo para ser distribuido o exportado a PDF.

---

## 🎨 Personalización de Diseño y Estilos

Puedes ajustar la paleta de colores corporativos directamente en las constantes iniciales de `cv.py`:

```python
# --- PALETA DE COLORES (Formato HEX) ---
COLOR_TEAL_OSCURO = "005B96"  # Color principal de cabecera y títulos
COLOR_GRIS_CLARO = "E8EEF2"   # Fondo de la barra lateral izquierda
COLOR_BLANCO = "FFFFFF"       # Fondo de la columna principal y textos claros
COLOR_TEXTO_OSCURO = "222222" # Color del texto principal de lectura
COLOR_MUTED = "555555"        # Subtítulos, fechas y notas
```

---

## 🔧 Detalles Técnicos y XML Nativo

`cv.py` utiliza manipulaciones nativas del estándar **WordprocessingML (OpenXML)** mediante `docx.oxml` para lograr funcionalidades avanzadas que no están soportadas nativamente por la API estándar de `python-docx`:

- **`set_cell_background(cell, color_hex)`**: Inyecta elementos `w:shd` para colorear fondos de celdas y bloques.
- **`set_cell_margins(cell, top, left, bottom, right)`**: Define padding interno mediante `w:tcMar` para evitar que el texto toque los límites de la celda.
- **`add_bottom_border(p, color_hex, size)`**: Crea líneas divisorias inferiores con `w:pBdr`.
- **`add_hyperlink(p, url, text, size, color_hex)`**: Genera relaciones externas de tipo hipervínculo en el paquete OPC de Word.
- **`allow_row_to_break(row)`**: Configura `w:cantSplit = 0` para permitir que el contenido fluya entre páginas sin cortar tablas.
- **`add_run(p, text, size, bold, color_hex, italic)`**: Escribe texto formateado garantizando que Word aplique el color HEX y tipografía.

---

## 🤝 Contribución

Las contribuciones son bienvenidas:

1. Haz un Fork del repositorio.
2. Crea tu rama (`git checkout -b feature/MejoraDiseno`).
3. Confirma tus cambios (`git commit -m 'feat: Añadida nueva plantilla de diseño'`).
4. Haz push a la rama (`git push origin feature/MejoraDiseno`).
5. Abre un **Pull Request**.

---

## 📝 Licencia

Este proyecto está bajo la Licencia [MIT](LICENSE).
