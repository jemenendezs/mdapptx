# 📄 Importar Markdown a PowerPoint

Este proyecto en VBA permite transformar un archivo en formato Markdown (con títulos y viñetas) en una presentación de PowerPoint. La macro interpreta títulos como nuevas diapositivas y viñetas como elementos de lista dentro de cada diapositiva.

## 🎯 Objetivo

Facilitar la creación rápida de presentaciones a partir de archivos `.md`, `.markdown` o `.txt` estructurados con:

- `#` o `##` para títulos.
- `-` para viñetas.

## 📁 Formato esperado del archivo Markdown

```markdown
# Título principal

## Sección 1
- Punto uno
- Punto dos

## Sección 2
- Otro punto

```

## ⚙️ ¿Qué hace el código?

1. Abre un cuadro de diálogo para seleccionar un archivo Markdown.
2. Lee el contenido del archivo, tratando de decodificarlo en UTF-8.
3. Procesa línea por línea:

   * Si es un **título**, crea una nueva diapositiva y lo coloca como título.
   * Si es una **viñeta**, la agrega como punto en la diapositiva actual.
4. Limpia referencias como `[^1]` y espacios innecesarios.
5. Agrega punto final a las viñetas si no terminan en `.`, `!`, `?` o `:`.
6. Muestra un resumen final con estadísticas de procesamiento.

## 🧩 Componentes del código

### `ImportarMarkdownAPowerPoint()`

* Macro principal que controla todo el flujo.
* Muestra un cuadro de diálogo para seleccionar un archivo Markdown.
* Lee el contenido del archivo como texto (preferiblemente en UTF-8).
* Divide el contenido en líneas individuales.
* Recorre cada línea:

  * Si es un título (`#` o `##`), crea una nueva diapositiva y establece el texto como título.
  * Si es una viñeta (`-`), la agrega como ítem a la diapositiva actual.
  * Muestra al final un mensaje con estadísticas del procesamiento (líneas, títulos, viñetas, diapositivas).

### `LeerArchivoUTF8(rutaArchivo)`

* Intenta leer el archivo usando `ADODB.Stream` con codificación UTF-8.
* Si falla, utiliza un método alternativo tradicional de lectura.
* Devuelve el contenido del archivo como cadena de texto.

### `LimpiarTexto(texto)`

* Elimina referencias como `[^1]`, `[^nota]`, etc.
* Reemplaza espacios dobles por espacios simples.
* Elimina espacios al inicio y al final de la cadena.
* Devuelve el texto limpio, listo para usarse como título o viñeta.

### `AgregarPuntoFinal(texto)`

* Se aplica únicamente a las viñetas.
* Agrega un punto final si el texto no termina en `.`, `!`, `?` o `:`.
* Devuelve el texto con puntuación corregida.

### `EsTitulo(linea)`

* Verifica si una línea comienza con `#`, `##` u otros signos de numeral.
* Permite formatos como:

  * `# Título`
  * `## Subtítulo`
  * `### Subsubtítulo`
* Devuelve `True` si la línea es un título válido.

### `ObtenerTitulo(linea)`

* Elimina los signos de numeral `#` del inicio.
* Limpia el texto eliminando referencias y espacios innecesarios.
* Si el título está vacío, lo reemplaza por `"Diapositiva N"` (donde N es el número correspondiente).
* Devuelve el texto limpio que se usará como título de la diapositiva.

### `EsVineta(linea)`

* Verifica si una línea comienza con `-` seguido de espacio.
* Devuelve `True` si es una viñeta válida.

### `ObtenerTextoVineta(linea)`

* Elimina el guion `-` inicial y los espacios.
* Limpia el texto de referencias y espacios.
* Aplica `AgregarPuntoFinal` para asegurar puntuación correcta.
* Si la viñeta está vacía, se reemplaza por `"Elemento de lista."`.
* Devuelve el texto limpio que se insertará como viñeta.

### `CrearNuevaDiapositiva(pres)`

* Añade una nueva diapositiva al final de la presentación.
* Usa el diseño predeterminado con título y contenido (`ppLayoutText`).
* Devuelve un objeto `Slide`.

### `AgregarTitulo(diap, textoTitulo)`

* Busca la forma correspondiente al título en la diapositiva.
* Si no la encuentra claramente, intenta usar la primera forma disponible.
* Inserta el texto del título en el marcador de posición.
* Maneja errores si ocurre algún problema al asignar el texto.

### `AgregarVineta(diap, textoVineta)`

* Busca el marcador de contenido (caja de texto principal).
* Agrega la viñeta al final del contenido existente (si lo hay).
* Aplica formato de lista con viñetas visuales.
* Si no encuentra forma adecuada, intenta usar la última forma de la diapositiva.
* También maneja errores y muestra mensajes si algo falla.

## 🛠️ Requisitos

* Microsoft PowerPoint (con macros habilitadas).
* Archivo de entrada en formato `.md`, `.markdown` o `.txt`.

## 💬 Mensaje final

Al finalizar el procesamiento, la macro muestra un resumen con:

* Total de líneas procesadas.
* Cantidad de títulos detectados.
* Cantidad de viñetas detectadas.
* Número de diapositivas creadas.

