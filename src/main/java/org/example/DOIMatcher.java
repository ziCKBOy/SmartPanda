package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class DOIMatcher {

    public static void main(String[] args) throws Exception {
        FileInputStream fis1 = new FileInputStream("C:\\Users\\Oliver Becerra\\OneDrive\\Desktop\\ayuda Vivi\\Scopus 2024.xlsx");
        FileInputStream fis2 = new FileInputStream("C:\\Users\\Oliver Becerra\\OneDrive\\Desktop\\ayuda Vivi\\Publicaciones indexadas actualizadas a marzo 2025.xlsx");

        Workbook wb1 = new XSSFWorkbook(fis1);
        Workbook wb2 = new XSSFWorkbook(fis2);

        Sheet sheet1 = wb1.getSheetAt(0);
        Sheet sheet2 = wb2.getSheet("Papers Indexados");

        if (sheet2 == null) {
            throw new RuntimeException("No se encontró la hoja 'Papers Indexados' en Planilla2.");
        }

        // Estilos
        CellStyle greenStyle = wb1.createCellStyle();
        greenStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle blueStyle = wb1.createCellStyle();
        blueStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        blueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Leer todos los DOI de la hoja 'Papers Indexados' en un Set para búsqueda rápida
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
            throw new RuntimeException("Columna 'INDEXACION' no encontrada en la hoja 'Papers Indexados'.");
        }

        for (Row row : sheet2) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String value = cell.getStringCellValue();
                    if (value.contains("10.")) {
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
                    String DOI = null;
                    //System.out.println(value);
                    String[] partes = value.split(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)", -1);

                    for (String parte : partes) {
                        String valor = parte.replace("\"", "").trim(); // Eliminar comillas y espacios

                        if (valor.startsWith("10.")) {
                            //System.out.println("DOI encontrado en planilla 1: " + valor+", buscando en planilla 2...");
                            DOI = valor;
                            break; // Salir si ya lo encontró
                        }
                    }


                        if (doiSet.contains(DOI)) {
                            System.out.println("DOI encontrado en planilla 2: " + DOI);
                            foundDOI = true;
                            matchedDOI = DOI;
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
                if (matchRow != null) {
                    Cell indexacionCell = matchRow.getCell(indexacionCol);
                    if (indexacionCell != null && indexacionCell.getCellType() == CellType.STRING) {
                        String current = indexacionCell.getStringCellValue();
                        //System.out.println("celda para completar:"+current);
                        if (!current.toUpperCase().contains("SCOPUS")) {
                            indexacionCell.setCellValue(current + "-SCOPUS");
                        }
                    }
                } else {
                    System.out.println("⚠️ DOI encontrado en planilla2, pero ya está marcado con SCOPUS: " + matchedDOI);
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
