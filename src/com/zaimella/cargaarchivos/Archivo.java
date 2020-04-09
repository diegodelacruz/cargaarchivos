/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.zaimella.cargaarchivos;

import static com.zaimella.cargaarchivos.CargaArchivos.recuperaConfiguracion;
import com.zaimella.constantes.Constantes;
import com.zaimella.excepciones.ArchivoConfiguraciones;
import com.zaimella.excepciones.InsertarRegistro;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.sql.Connection;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Diego de la Cruz <www.zaimella.com>
 */
public class Archivo {

    public static void listarArchivos(String ruta, List<File> archivos) {
        // Ruta de los archivos
        File directorio = new File(ruta);

        // Obtengo todos los archivos
        File[] listaArchivos = directorio.listFiles();

        if (listaArchivos != null) {
            for (File archivo : listaArchivos) {
                if (archivo.isFile()) {
                    archivos.add(archivo);
                } else if (archivo.isDirectory()) {
                    listarArchivos(archivo.getAbsolutePath(), archivos);
                }
            }
        }
    }

    public static void leerArchivo(File archivo) throws IOException, FileNotFoundException, InvalidFormatException {
        String extension = FilenameUtils.getExtension(archivo.getName()).toLowerCase();
        String pad = "..........................................................................................";
        String nme = "Archivo " + archivo.getName();

        if (extension.equals("txt") || extension.equals("csv")) {
            System.out.print(nme + pad.substring(nme.length()));
            archivoPlano(archivo);
        } else if (extension.equals("xlsx") || extension.equals("xls")) {
            System.out.print(nme + pad.substring(nme.length()));
            archivoOffice(archivo);
        }
    }

    public static void archivoPlano(File archivo) {
        BufferedReader br = null;
        String vacio = "";
        String separador = "<-?->";
        String reg;
        String pad;
        int lim = 5000;
        int i = 0;

        Connection conexion = null;
        try {
            Properties properties = recuperaConfiguracion();

            String host = properties.getProperty("dbHost");
            String puerto = properties.getProperty("dbPuerto");
            String servicio = properties.getProperty("dbNombreServicio");
            String usuario = properties.getProperty("dbUsuario");
            String password = properties.getProperty("dbClave");

            ZaimellaDB zaimellaDB = new ZaimellaDB();
            conexion = zaimellaDB.getConnectionDB(host, puerto, servicio, usuario, password);
            conexion.setAutoCommit(false);
            Statement stmt = conexion.createStatement();

            try {
                br = new BufferedReader(new FileReader(archivo.getAbsolutePath()));
            } catch (FileNotFoundException fnfe) {
                System.out.println("Archivo no encontrado." + fnfe);
            }

            try {
                String linea;
                while ((linea = br.readLine()) != null) {
                    i++;
                    linea = linea.replace("\"", "").replaceAll(",", "").replaceAll("\\'", "").replaceAll("\\t", separador);

                    zaimellaDB.insertaTablaCargaArchivos(conexion, archivo.getParent(), archivo.getName(), FilenameUtils.getExtension(archivo.getName()).toLowerCase(), vacio, Integer.valueOf(i), null, linea);

                    if (i % lim == 0) {
                        conexion.commit();
                    }

                }
                conexion.commit();
                reg = Integer.toString(i);
                pad = "..........";
                System.out.println(pad.substring(reg.length()) + i + " registros cargados");

            } catch (IOException ioe) {
                System.out.println("Error al leer el archivo. " + ioe.getMessage());
            }

        } catch (ArchivoConfiguraciones e) {
            System.out.println("Archivo de configuración no encontrado, el path buscado es : " + Constantes.ARCHIVO_CONFIGURACION);
        } catch (Exception e) {
            System.out.println(e);
        } finally {
            ZaimellaDB.cerrarConexion(conexion);
        }
    }

    public static void archivoOffice(File archivo) throws FileNotFoundException, IOException, InvalidFormatException {
        String extension = FilenameUtils.getExtension(archivo.getName()).toLowerCase();
        FileInputStream fis = new FileInputStream(archivo.getAbsolutePath());
        String nombreHoja;
        String reg;
        String pad;
        String nhj;
        String separador = "<-?->";
        DataFormatter formatoFecha = new DataFormatter();
        List<String> lineasArchivo = new ArrayList();
        int lim = 5000;
        int primeraFila;
        int ultimaFila;
        int ultimaColumna;
        int k;

        Connection conexion = null;
        try {
            Properties properties = recuperaConfiguracion();

            String host = properties.getProperty("dbHost");
            String puerto = properties.getProperty("dbPuerto");
            String servicio = properties.getProperty("dbNombreServicio");
            String usuario = properties.getProperty("dbUsuario");
            String password = properties.getProperty("dbClave");

            ZaimellaDB zaimellaDB = new ZaimellaDB();
            conexion = zaimellaDB.getConnectionDB(host, puerto, servicio, usuario, password);
            conexion.setAutoCommit(false);
            Statement stmt = conexion.createStatement();

            if (extension.equals("xls")) {
                HSSFWorkbook libro = new HSSFWorkbook(fis);
                for (int i = 0; i < libro.getNumberOfSheets(); i++) {
                    nombreHoja = libro.getSheetName(i).trim();
                    HSSFSheet hoja = libro.getSheetAt(i);

                    pad = "...............................";
                    nhj = "  Hoja " + nombreHoja;
                    System.out.print(nhj + pad.substring(nhj.length()));

                    primeraFila = Math.min(0, hoja.getFirstRowNum());
                    ultimaFila = Math.max(0, hoja.getLastRowNum());

                    k = 0;
                    for (int fila = primeraFila; fila <= ultimaFila; fila++) {
                        k++;
                        Row row = hoja.getRow(fila);
//                        int ultimaColumna = Math.max(0, 30);
                        ultimaColumna = row.getLastCellNum();
                        StringBuilder sb = new StringBuilder();

                        if (row != null) {
                            for (int columna = 0; columna <= (ultimaColumna - 1); columna++) {
                                Cell cell = row.getCell(columna, Row.RETURN_BLANK_AS_NULL);
                                CellType type;
                                try {
                                    type = cell.getCellTypeEnum();
                                } catch (NullPointerException npe) {
                                    type = CellType.BLANK;
                                }

                                String texto = formatoFecha.formatCellValue(cell);

                                if (type == CellType.STRING) {
                                    sb.append(cell.getRichStringCellValue().toString()).append("\t");
                                } else if (type == CellType.NUMERIC) {
                                    cell.setCellType(CellType.STRING);
                                    sb.append(cell.getRichStringCellValue().toString()).append("\t");
                                } else if (type == CellType.BOOLEAN) {
                                    cell.setCellType(CellType.STRING);
                                    sb.append(cell.getRichStringCellValue().toString()).append("\t");
                                } else if (type == CellType.FORMULA) {
                                    if (cell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
                                        cell.setCellType(CellType.STRING);
                                        sb.append(cell.getRichStringCellValue().toString()).append("\t");
                                    } else if (cell.getCachedFormulaResultTypeEnum() == CellType.STRING) {
                                        sb.append(cell.getRichStringCellValue().toString()).append("\t");
                                    } else if (cell.getCachedFormulaResultTypeEnum() == CellType.ERROR) {
                                        sb.append("[orcl]").append("\t");
                                    }
                                } else if (type == CellType.BLANK) {
                                    sb.append("[orcl]").append("\t");
                                } else if (type == CellType.ERROR) {
                                    sb.append("[orcl]").append("\t");
                                } else if (cell == null) {
                                }
                            }

                            String linea = new String();
                            linea = sb.toString();
                            linea = linea.replace("\"", "").replaceAll(",", ".").replaceAll("\\'", "").replaceAll("\\t", separador);

                            zaimellaDB.insertaTablaCargaArchivos(conexion, archivo.getParent(), archivo.getName(), extension, nombreHoja, Integer.valueOf(k), Integer.valueOf(ultimaColumna), linea);

                            if (k % lim == 0) {
                                conexion.commit();
                            }
                        }
                    }
                    conexion.commit();
                    reg = Integer.toString(k);
                    pad = "..........";
                    System.out.println(pad.substring(reg.length()) + k + " registros cargados");
                }
            } else if (extension.equals("xlsx")) {
                XSSFWorkbook libro = new XSSFWorkbook(fis);
                for (int i = 0; i < libro.getNumberOfSheets(); i++) {
                    nombreHoja = libro.getSheetName(i).trim();
                    XSSFSheet hoja = libro.getSheetAt(i);

                    pad = "...............................";
                    nhj = "  Hoja " + nombreHoja;
                    System.out.print(nhj + pad.substring(nhj.length()));

                    primeraFila = Math.min(0, hoja.getFirstRowNum());
                    ultimaFila = Math.max(0, hoja.getLastRowNum());

                    k = 0;
                    for (int fila = primeraFila; fila <= ultimaFila; fila++) {
                        k++;
                        Row row = hoja.getRow(fila);
//                        int ultimaColumna = Math.max(0, 30);
                        ultimaColumna = row.getLastCellNum();
                        StringBuilder sb = new StringBuilder();

                        if (row != null) {
                            for (int columna = 0; columna <= (ultimaColumna - 1); columna++) {
                                Cell cell = row.getCell(columna, Row.RETURN_BLANK_AS_NULL);
                                CellType type;
                                try {
                                    type = cell.getCellTypeEnum();
                                } catch (NullPointerException e) {
                                    type = CellType.BLANK;
                                }

                                String text = formatoFecha.formatCellValue(cell);

                                if (type == CellType.STRING) {
                                    sb.append(cell.getRichStringCellValue().toString()).append("\t");
                                } else if (type == CellType.NUMERIC) {
                                    cell.setCellType(CellType.STRING);
                                    sb.append(cell.getRichStringCellValue()).append("\t");
                                } else if (type == CellType.BOOLEAN) {
                                    cell.setCellType(CellType.STRING);
                                    sb.append(cell.getRichStringCellValue()).append("\t");
                                } else if (type == CellType.FORMULA) {
                                    if (cell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
                                        cell.setCellType(CellType.STRING);
                                        sb.append(cell.getRichStringCellValue().toString()).append("\t");
                                    } else if (cell.getCachedFormulaResultTypeEnum() == CellType.STRING) {
                                        sb.append(cell.getRichStringCellValue().toString()).append("\t");
                                    } else if (cell.getCachedFormulaResultTypeEnum() == CellType.ERROR) {
                                        sb.append("[orcl]").append("\t");
                                    }
                                } else if (type == CellType.BLANK) {
                                    sb.append("[orcl]").append("\t");
                                } else if (type == CellType.ERROR) {
                                    sb.append("[orcl]").append("\t");
                                } else if (cell == null) {
                                }
                            }

                            String linea = new String();
                            linea = sb.toString();
                            linea = linea.replace("\"", "").replaceAll(",", ".").replaceAll("\\'", "").replaceAll("\\t", separador);

                            zaimellaDB.insertaTablaCargaArchivos(conexion, archivo.getParent(), archivo.getName(), extension, nombreHoja, Integer.valueOf(k), Integer.valueOf(ultimaColumna), linea);

                            if (k % lim == 0) {
                                conexion.commit();
                            }
                        }
                    }
                    conexion.commit();
                    reg = Integer.toString(k);
                    pad = "..........";
                    System.out.println(pad.substring(reg.length()) + k + " registros cargados");
                }
            }
        } catch (ArchivoConfiguraciones e) {
            System.out.println("Archivo de configuración no encontrado, el path buscado es : " + Constantes.ARCHIVO_CONFIGURACION);
        } catch (Exception e) {
            System.out.println(e);
        } finally {
            ZaimellaDB.cerrarConexion(conexion);
        }
    }
}
