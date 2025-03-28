package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import lombok.extern.log4j.Log4j2;

@Log4j2
public class Main {
    public static void main(String[] args) throws Exception{


        FileInputStream fis1 = new FileInputStream("C:\\Users\\Oliver Becerra\\OneDrive\\Desktop\\ayuda Vivi\\Scopus 2024.xlsx");
        FileInputStream fis2 = new FileInputStream("C:\\Users\\Oliver Becerra\\OneDrive\\Desktop\\ayuda Vivi\\Publicaciones indexadas actualizadas a marzo 2025.xlsx");

        Workbook wb1 = new XSSFWorkbook(fis1);
        Workbook wb2 = new XSSFWorkbook(fis2);

        Sheet sheet1 = wb1.getSheetAt(0);
        Sheet sheet2 = wb2.getSheetAt(0);

        // Estilos
        CellStyle greenStyle = wb1.createCellStyle();
        greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle blueStyle = wb1.createCellStyle();
        blueStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        blueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Leer todos los DOI de la segunda planilla en un Set para búsqueda rápida
        Set<String> doiSet = new HashSet<>();
        Map<String, Row> doiToRowMap = new HashMap<>();
        int indexacionCol = -1;

        Row headerRow = sheet2.getRow(0);
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase("INDEXACION")) {
                indexacionCol = cell.getColumnIndex();
                break;
            }
        }

        if (indexacionCol == -1) {
            log.info("Columna 'INDEXACION' no encontrada en Planilla2");
            throw new RuntimeException("Columna 'INDEXACION' no encontrada en Planilla2");
        }

        for (Row row : sheet2) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String value = cell.getStringCellValue();
                    if (value.startsWith("10.")) {
                        doiSet.add(value);
                        doiToRowMap.put(value, row);
                    }
                }
            }
        }

        // Procesar Planilla1
        for (Row row : sheet1) {
            boolean foundDOI = false;
            String matchedDOI = null;

            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String value = cell.getStringCellValue();
                    if (value.startsWith("10.") && doiSet.contains(value)) {
                        log.info("DOI encontrado");
                        foundDOI = true;
                        matchedDOI = value;
                        break;
                    }
                }
            }

            CellStyle styleToApply = foundDOI ? greenStyle : blueStyle;

            for (Cell cell : row) {
                cell.setCellStyle(styleToApply);
            }

            // Concatenar INDEXACION en planilla 2 si DOI encontrado
            if (foundDOI && matchedDOI != null) {
                Row matchRow = doiToRowMap.get(matchedDOI);
                Cell indexacionCell = matchRow.getCell(indexacionCol);
                if (indexacionCell != null && indexacionCell.getCellType() == CellType.STRING) {
                    String current = indexacionCell.getStringCellValue();
                    if (!current.contains("Scopus")) {
                        log.info("concatenando Scopus en 2da planilla");
                        indexacionCell.setCellValue(current + "-ScopusIA");
                    }
                }
            }
        }

        fis1.close();
        fis2.close();

        FileOutputStream fos1 = new FileOutputStream("C:\\Users\\Oliver Becerra\\OneDrive\\Desktop\\ayuda Vivi\\Planilla1_coloreada.xlsx");
        wb1.write(fos1);
        fos1.close();

        FileOutputStream fos2 = new FileOutputStream("C:\\Users\\Oliver Becerra\\OneDrive\\Desktop\\ayuda Vivi\\Planilla2_actualizada.xlsx");
        wb2.write(fos2);
        fos2.close();

        wb1.close();
        wb2.close();

        System.out.println("Proceso completado. Archivos guardados.");

    }
}