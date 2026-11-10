<div align="center">

# 🌱 Evaluación Nutricional Infantil — OMS / SISVAN

**Una sola planilla (sin macros) que calcula los puntajes-z de la OMS y el diagnóstico nutricional de niños y adolescentes (0 a 19 años) — para apoyar el seguimiento del crecimiento en guarderías, escuelas y servicios de salud.**

🌐 **Sitio oficial: [isasaade-23.github.io/avaliacao-nutricional-oms](https://isasaade-23.github.io/avaliacao-nutricional-oms/)** (versiones Excel, LibreOffice y Google Sheets)

[🇧🇷 Português](README.md) · [🇺🇸 English](README.en.md) · 🇪🇸 **Español**

[![Licencia CC BY-NC-SA 4.0](https://img.shields.io/badge/licencia-CC%20BY--NC--SA%204.0-green)](https://creativecommons.org/licenses/by-nc-sa/4.0/)
![Hecho con Python](https://img.shields.io/badge/hecho%20con-Python-blue)
![Estándar OMS](https://img.shields.io/badge/est%C3%A1ndar-WHO%20Anthro%20%2F%20AnthroPlus-orange)
![Sin macros](https://img.shields.io/badge/Excel-sin%20macros-lightgrey)
[![Sponsor](https://img.shields.io/badge/apoyar-GitHub%20Sponsors-EA4AAA?logo=github-sponsors)](https://github.com/sponsors/isasaade-23)
[![Ko-fi](https://img.shields.io/badge/apoyar-Ko--fi-FF5E5B?logo=ko-fi&logoColor=white)](https://ko-fi.com/ailuciola)

</div>

---

## 🎯 Para qué sirve

Seguir el **estado nutricional de un niño a lo largo del tiempo** suele ser tedioso: exige consultar las tablas de la OMS, calcular los puntajes-z a mano y clasificarlos según las guías nacionales (el SISVAN, en Brasil). En guarderías, escuelas y unidades de salud muchas veces no se hace — o se hace con errores.

Este proyecto ofrece una herramienta **simple y confiable** para quien cuida niños:

- Se **escriben peso, talla y fechas** en la pestaña **Banco**;
- La **misma planilla devuelve automáticamente** los puntajes-z y el diagnóstico nutricional, con colores;
- Se puede **seguir al mismo niño mes a mes** (crecimiento longitudinal) y ver la curva de evolución.

El objetivo es **mejorar la vigilancia nutricional** — detectar temprano la delgadez, la baja talla, el sobrepeso y la obesidad — en servicios que atienden niños pero rara vez cuentan con un sistema para ello.

> **Sin macros. En las versiones Excel y LibreOffice ningún dato sale de tu computadora.**

---

## ✨ Características

- 📏 **Cuatro índices antropométricos** de la OMS: IMC/edad, Talla/edad, Peso/edad y Peso/talla.
- 🎨 **Diagnóstico automático** con las categorías del SISVAN / Ministerio de Salud, con colores (verde = adecuado, amarillo = atención, rojo = déficit/exceso).
- 📈 **Seguimiento longitudinal**: elige un niño y observa el historial y la curva de puntajes-z.
- 👥 **Panel del grupo/unidad**: distribución del estado nutricional del grupo.
- 📄 **Un solo archivo**: entrada de datos y cálculo en la misma planilla, sin nada que enlazar.
- ✅ **Validado**: las fórmulas de Excel se contrastaron con una implementación de referencia en Python y con el paquete R oficial de la OMS.

---

## 🚀 Cómo usarlo

La calculadora es un **archivo único**, en tres versiones con el mismo contenido. Elige la que corresponde a tu programa:

| Versión | Archivo | Descarga |
|---|---|---|
| Excel | `excel_Calculadora_OMS.xlsx` | [descargar](https://github.com/isasaade-23/avaliacao-nutricional-oms/raw/main/excel_Calculadora_OMS.xlsx) |
| LibreOffice | `libreoffice_Calculadora_OMS.ods` | [descargar](https://github.com/isasaade-23/avaliacao-nutricional-oms/raw/main/libreoffice_Calculadora_OMS.ods) |
| Google Sheets | `gsheet_Calculadora_OMS` | [crear mi copia](https://docs.google.com/spreadsheets/d/1ZrYO0hEXhPV054k1nwWsasHnFxFsZ1CQfY2td47YB90/copy) · [ver sin iniciar sesión](https://docs.google.com/spreadsheets/d/1ZrYO0hEXhPV054k1nwWsasHnFxFsZ1CQfY2td47YB90/preview) |

Puedes ver el modelo de Google Sheets sin iniciar sesión. Para escribir y guardar datos, elige "crear mi copia" e inicia sesión en una cuenta de Google, si se solicita. Confirma con "Hacer una copia": se guarda en tu propio Drive, y ahí escribes los datos. No necesitas solicitar permiso para editar el modelo original. Para usar la calculadora sin una cuenta de Google, descarga la versión Excel o LibreOffice con los enlaces de arriba.

1. **Escribe las mediciones** en la pestaña **Banco** — una fila por medición. Para seguir a un niño, reutiliza siempre el **mismo ID** y agrega una fila en cada visita.
2. **Lee los resultados** en la pestaña **Cálculo**, con colores. Usa **Ficha** para el historial de un niño y **Painel** para la vista del grupo. La pestaña **Instruções** trae el paso a paso dentro de la propia planilla.

Los datos y los cálculos están en el mismo archivo, así que no hay nada que enlazar. La guía de uso incluida (`.pdf`/`.docx`, en portugués) lo explica en lenguaje sencillo.

---

## 📊 Índices y clasificación (SISVAN / Ministerio de Salud)

| Índice | Rango de edad | Categorías (por puntaje-z) |
|---|---|---|
| **IMC/edad** (< 5 años) | 0–60 meses | <−3 Delgadez severa · −3 a <−2 Delgadez · −2 a +1 Normal · >+1 a +2 Riesgo de sobrepeso · >+2 a +3 Sobrepeso · >+3 Obesidad |
| **IMC/edad** (5–19 años) | 61–228 meses | <−3 Delgadez severa · −3 a <−2 Delgadez · −2 a +1 Normal · >+1 a +2 Sobrepeso · >+2 a +3 Obesidad · >+3 Obesidad grave |
| **Talla/edad** | 0–228 meses | <−3 Talla muy baja · −3 a <−2 Talla baja · ≥−2 Adecuada |
| **Peso/edad** | 0–120 meses (≤ 10 a) | <−3 Muy bajo · −3 a <−2 Bajo · −2 a +2 Adecuado · >+2 Elevado |
| **Peso/talla** | 0–60 meses (≤ 5 a) | mismos cortes que IMC/edad < 5 años |

Fuera de estos rangos el resultado es "—" (Peso/edad no existe por encima de 10 años y Peso/talla por encima de 5 años en el estándar de la OMS).

---

## 🔬 Base metodológica

Los puntajes-z siguen el **método LMS**, oficial de la OMS:

$$z = \frac{(X/M)^L - 1}{L \cdot S} \quad \text{(o } z = \frac{\ln(X/M)}{S} \text{ cuando } L = 0\text{)}$$

donde **L, M, S** son parámetros por sexo y edad (o longitud/talla), tomados de las tablas oficiales de la OMS.

Detalles que reproducen fielmente WHO Anthro / AnthroPlus:

- **Edad:** de 0 a 5 años se calcula **en días** (WHO Anthro 2006); de 5 a 19 años, **en meses** (WHO Growth Reference 2007).
- **Ajuste de la OMS para |z| > 3** en los índices basados en peso (IMC/edad, Peso/edad, Peso/talla), evitando extrapolar más allá de los datos observados.
- **Ajuste acostado/de pie (±0,7 cm):** aplicado cuando la forma de medir difiere del estándar para la edad (< 2 años acostado; ≥ 2 años de pie).
- **Solo funciones clásicas de Excel** (INDICE/COINCIDIR/SI) → funciona también en versiones antiguas, sin macros.

### Fuentes de los datos
- **WHO Child Growth Standards (2006)** — 0 a 5 años — paquete oficial [`anthro`](https://github.com/worldhealthorganization/anthro).
- **WHO Growth Reference (2007)** — 5 a 19 años — paquete oficial [`anthroplus`](https://github.com/worldhealthorganization/anthroplus).

Los archivos fuente originales de la OMS se conservan por separado para preservar su procedencia.

### Validación
- La lógica de referencia se validó contra el paquete **R** oficial de la OMS en miles de casos.
- **16 casos de prueba** se recalcularon en **Excel real** y coinciden con la referencia, incluidos los extremos (|z|>3) y los límites de edad (0, 24, 60, 61, 120 y 228 meses).

---

## 🗂️ Estructura del proyecto

| Archivo | Descripción |
|---|---|
| `excel_Calculadora_OMS.xlsx` | **Entregable** — la calculadora para Excel (tablas de la OMS + fórmulas + entrada de datos) |
| `libreoffice_Calculadora_OMS.ods` | **Entregable** — la misma calculadora para LibreOffice |
| `Guia rápido — Avaliação Nutricional.docx` | Guía de uso en lenguaje sencillo (en portugués) |
| `Guia rápido — Avaliação Nutricional.pdf` | Guía imprimible (en portugués) |
| `docs/index.html` | Página oficial de descargas |

Los scripts de construcción y validación no se incluyen en este repositorio. Para solicitar información que permita auditar los cálculos, utiliza el correo de contacto de abajo.

---

## 📖 Cómo citar

Si este proyecto te resulta útil, una cita es muy bienvenida 🙏 (mira el botón **"Cite this repository"** en GitHub, o el archivo [`CITATION.cff`](CITATION.cff)):

> Saade, I. (2026). *Evaluación Nutricional Infantil — OMS / SISVAN* [software]. https://github.com/isasaade-23/avaliacao-nutricional-oms

Si quieres apoyar el mantenimiento del proyecto, es totalmente opcional — vía [GitHub Sponsors](https://github.com/sponsors/isasaade-23), [Ko-fi](https://ko-fi.com/ailuciola) o Pix (Brasil, clave: isasaade23@gmail.com).

---

## 🔒 Privacidad

Son **datos de salud de menores**. Usa el **ID/número de registro** como clave, mantén el **Nombre** como opcional y guarda el archivo en un **lugar restringido y controlado**. Las versiones Excel y LibreOffice funcionan **completamente sin conexión**. En la versión Google Sheets los datos quedan en la nube de Google: haz una copia en tu Drive, no escribas nada en el modelo original y restringe el acceso a tu copia.

---

## 📄 Licencia

Distribuido bajo **[CC BY-NC-SA 4.0](LICENSE)**: puedes **usar, copiar, modificar y distribuir** libremente, incluso en guarderías, escuelas y servicios de salud. Las condiciones son: **conservar el aviso de autoría**, **no venderlo** (ni el original ni versiones modificadas) y, si lo adaptas, **redistribuirlo bajo la misma licencia**. Si puedes, **cita el proyecto** — me ayuda mucho. 💚
