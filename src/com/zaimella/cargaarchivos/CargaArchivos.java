/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.zaimella.cargaarchivos;

import com.zaimella.constantes.Constantes;
import com.zaimella.excepciones.ArchivoConfiguraciones;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
/**
 *
 * @author Diego de la Cruz <www.zaimella.com>
 */
public class CargaArchivos {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, FileNotFoundException, InvalidFormatException {
        BufferedReader br = null;

        String directorio = "D:\\CargasMasivas\\";
        List<File> listaArchivos = new ArrayList();

        Archivo.listarArchivos(directorio, listaArchivos);
                
        for (int i = 0; i < listaArchivos.size(); i++) {
            File file = listaArchivos.get(i);
            Archivo.leerArchivo(file);            
            file = null;
        }
    }

    public static Properties recuperaConfiguracion()
            throws Exception {
        Properties properties = new Properties();

        File file = new File(Constantes.ARCHIVO_CONFIGURACION);
        if (file.exists()) {
            properties.load(FileUtils.openInputStream(file));
            return properties;
        }
        throw new ArchivoConfiguraciones();
    }
}
