package org.example;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        // Crea una lista para almacenar los registros
            List<Registro> registros = new ArrayList<>();
      try{  
            File file = new File("Ejemplo.xlsx");

            String ruta = file.getAbsolutePath();

            System.out.println("ruta: "+ruta);
            
            // Crea una instancia de la clase Workbook
            Workbook workbook = WorkbookFactory.create(new FileInputStream(ruta));

            /// Obtiene una referencia a la hoja de trabajo
            Sheet sheet = workbook.getSheetAt(0);

            

            // Itera sobre las filas de la hoja
            for (Row row : sheet) {
                Cell agenciaCell = row.getCell(0);
                Cell intermediarioCell = row.getCell(1);
                Cell ramoCell = row.getCell(2);
                Cell añoCell = row.getCell(3);
                Cell primaNetaCell = row.getCell(4);
                Cell detalleAnexoCell = row.getCell(5);
                Cell comisionCell = row.getCell(6);

                String agencia = (agenciaCell != null) ? agenciaCell.getStringCellValue() : "";
                String intermediario = (intermediarioCell != null) ? intermediarioCell.getStringCellValue() : "";
                String ramo = (ramoCell != null) ? ramoCell.getStringCellValue() : "";
                int año = 0;
                double primaNeta = 0;
                String detalleAnexo = (detalleAnexoCell != null) ? detalleAnexoCell.getStringCellValue() : "";
                double comision = 0;

                // Verifica el tipo de celda antes de obtener el valor
                if (añoCell != null && añoCell.getCellType() == CellType.NUMERIC) {
                    año = (int) añoCell.getNumericCellValue();
                }
                if (primaNetaCell != null && primaNetaCell.getCellType() == CellType.NUMERIC) {
                    primaNeta = primaNetaCell.getNumericCellValue();
                }
                if (comisionCell != null && comisionCell.getCellType() == CellType.NUMERIC) {
                    comision = comisionCell.getNumericCellValue();
                }

                // Crea una instancia de la clase Registro
                Registro registro = new Registro(agencia, intermediario, ramo, año, primaNeta, detalleAnexo, comision);
                registros.add(registro);
                
            }

            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

             // Realiza los filtros que necesitas
            List<Registro> livianos = new ArrayList<>();
            for (Registro registro : registros) {
                if (registro.getRamo().equals("autos livia")) {
                    livianos.add(registro);
                }
            }

            double maxPrimaNeta = 0;
            List<Registro> mayorPrima = new ArrayList<>();
            for (Registro registro : registros) {
                if (registro.getPrimaNeta() > maxPrimaNeta) {
                    maxPrimaNeta = registro.getPrimaNeta();
                    mayorPrima.add(registro);
                }
            }

            List<Registro> expediciones = new ArrayList<>();
            for (Registro registro : registros) {
                if (registro.getDetalleAnexo().equals("EPEDICION")) {
                    expediciones.add(registro);
                }
            }

            double maxComision = 0;
            List<Registro> mayorComision = new ArrayList<>();
            for (Registro registro : registros) {
                if (registro.getComision() > maxComision) {
                    maxComision = registro.getComision();
                    mayorComision.add(registro);
                }
            }

            // Imprimir resultados
            System.out.println("Registros de autos livianos:");
            for (Registro registro : livianos) {
                System.out.println(registro.toString());
            }

            System.out.println("\nMayor prima neta:");
            System.out.println(mayorPrima.toString());
            for (Registro registro : mayorPrima) {
                System.out.println(registro.toString());
            }

            System.out.println("\nRegistros de expedición:");
            for (Registro registro : expediciones) {
                System.out.println(registro.toString());
            }

            System.out.println("\nMayor comisión:");
            for (Registro registro : mayorComision) {
                System.out.println(registro.toString());
            }
            
            
            escribirResultadosEnExcel("resultados.xlsx", livianos, mayorPrima, expediciones, mayorComision);
            
    }

    public static void escribirResultadosEnExcel(String nombreArchivo, List<Registro> livianos,
            List<Registro> mayorPrima, List<Registro> expediciones, List<Registro> mayorComision) {
                Row encabezadoRow = null;
                Cell encabezadoCell = null;
        try (Workbook workbook = new XSSFWorkbook()) {
            // Crea una hoja de trabajo en el libro de Excel
            Sheet sheet = workbook.createSheet("Resultados");

            // Crea la fila de encabezados de columna
            Row headerRow = sheet.createRow(0);
            String[] columnas = {
                "agencia solitaria", "principales intermediarios", "ramo", "año", "PRIMA NETA", "DETALLE DE ANEXO", "% COMISION"
            };
            // Escribe los encabezados de columna

            for (int i = 0; i < columnas.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columnas[i]);
            }

            // Escribe los resultados en las filas correspondientes
            int rowNum = 1;
            /* 
            ultimaFila = sheet.getLastRowNum();
            filaEncabezado = ultimaFila + 2;
            encabezadoRow = sheet.createRow(filaEncabezado);
            encabezadoCell = encabezadoRow.createCell(0);
            encabezadoCell.setCellValue("Registros de autos livianos:"); 
            */
            
            for (Registro registro : livianos) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(registro.getAgencia());
                row.createCell(1).setCellValue(registro.getIntermediario());
                row.createCell(2).setCellValue(registro.getRamo());
                row.createCell(3).setCellValue(registro.getAño());
                row.createCell(4).setCellValue(registro.getPrimaNeta());
                row.createCell(5).setCellValue(registro.getDetalleAnexo());
                row.createCell(6).setCellValue(registro.getComision());
                
            }

            rowNum++;
            encabezadoRow = sheet.createRow(rowNum);
            encabezadoCell = encabezadoRow.createCell(0);
            encabezadoCell.setCellValue("Mayor comisión:");

            rowNum = rowNum + 2;

            for (Registro registro : mayorComision) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(registro.getAgencia());
                row.createCell(1).setCellValue(registro.getIntermediario());
                row.createCell(2).setCellValue(registro.getRamo());
                row.createCell(3).setCellValue(registro.getAño());
                row.createCell(4).setCellValue(registro.getPrimaNeta());
                row.createCell(5).setCellValue(registro.getDetalleAnexo());
                row.createCell(6).setCellValue(registro.getComision());
                
            }
            rowNum++;

            encabezadoRow = sheet.createRow(rowNum);
            encabezadoCell = encabezadoRow.createCell(0);
            encabezadoCell.setCellValue("Registros de expedición:");

            rowNum = rowNum + 2;

            for (Registro registro : expediciones) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(registro.getAgencia());
                row.createCell(1).setCellValue(registro.getIntermediario());
                row.createCell(2).setCellValue(registro.getRamo());
                row.createCell(3).setCellValue(registro.getAño());
                row.createCell(4).setCellValue(registro.getPrimaNeta());
                row.createCell(5).setCellValue(registro.getDetalleAnexo());
                row.createCell(6).setCellValue(registro.getComision());
                
            }
            rowNum++;

            encabezadoRow = sheet.createRow(rowNum);
            encabezadoCell = encabezadoRow.createCell(0);
            encabezadoCell.setCellValue("Mayor comisión:");

            rowNum = rowNum + 2;

            for (Registro registro : mayorComision) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(registro.getAgencia());
                row.createCell(1).setCellValue(registro.getIntermediario());
                row.createCell(2).setCellValue(registro.getRamo());
                row.createCell(3).setCellValue(registro.getAño());
                row.createCell(4).setCellValue(registro.getPrimaNeta());
                row.createCell(5).setCellValue(registro.getDetalleAnexo());
                row.createCell(6).setCellValue(registro.getComision());
                
            }

            System.out.println("rowNum: "+rowNum);

            // Escribe los resultados adicionales si es necesario
            // ...

            // Guarda el libro de Excel en un archivo
            try (FileOutputStream outputStream = new FileOutputStream(nombreArchivo)) {
                workbook.write(outputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}