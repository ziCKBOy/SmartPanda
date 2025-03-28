# 🐼 Smart Panda

**Smart Panda** es una utilidad en Java para comparar dos archivos Excel (`.xlsx`) utilizando un código identificador (por ejemplo, un DOI que comienza con `10.`). 

El programa permite:

- ✅ Buscar coincidencias entre los dos archivos
- ✅ Marcar visualmente las coincidencias en el primer archivo
- ✅ Agregar una etiqueta de indexación en el segundo archivo

---

## 🚀 Funcionalidad

1. **Lectura de archivos Excel**
   - Lee dos planillas Excel: `Planilla1.xlsx` y `Planilla2.xlsx`.
   - En la segunda planilla, analiza exclusivamente la hoja llamada `Papers Indexados`.

2. **Búsqueda de identificadores**
   - Recorre las filas de la primera planilla en busca de valores que comiencen con `10.` (por ejemplo, DOI).
   - Compara estos valores con los contenidos en la segunda planilla.

3. **Coloreado de filas**
   - Si encuentra una coincidencia en la segunda planilla, **colorea la celda del identificador en verde**.
   - Si no la encuentra, **la colorea en azul**.

4. **Modificación de contenido**
   - Si el identificador existe en la segunda planilla, agrega el texto `-Scopus` al final del contenido de la columna `INDEXACION`, si aún no está presente.

---

## 📂 Archivos generados

- `Planilla1_coloreada.xlsx`: Primer archivo con celdas coloreadas según coincidencias.
- `Planilla2_actualizada.xlsx`: Segundo archivo con la columna `INDEXACION` actualizada.

---

## 🧪 Requisitos

- Java 17+
- Apache POI (biblioteca para manipular archivos Excel)

Puedes agregarlo con Maven:

```xml
<dependency>
  <groupId>org.apache.poi</groupId>
  <artifactId>poi-ooxml</artifactId>
  <version>5.2.3</version>
</dependency>
