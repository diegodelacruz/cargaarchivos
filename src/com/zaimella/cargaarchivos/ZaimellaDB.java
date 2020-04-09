/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.zaimella.cargaarchivos;

import com.zaimella.excepciones.InsertarRegistro;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;

/**
 *
 * @author Diego de la Cruz <www.zaimella.com>
 */
public class ZaimellaDB {

    public Connection getConnectionDB(String ipHost, String puerto, String sid, String username, String password)
            throws Exception {
//        System.out.println(".....Zaimella del Ecuador SA.....");

        Class.forName("oracle.jdbc.driver.OracleDriver");

        StringBuilder urlJdbc = new StringBuilder();
        urlJdbc.append("jdbc:oracle:thin:@").append(ipHost);
        urlJdbc.append(":");
        urlJdbc.append(puerto);
        urlJdbc.append(":");
        urlJdbc.append(sid);

        Connection conn = DriverManager.getConnection(urlJdbc.toString(), username, password);

//        System.out.println("¡Conectado!");
        return conn;
    }

    public void insertaTablaCargaArchivos(Connection conexion, String ruta, String archivo, String extension, String hoja, Integer linea, Integer separadores, String contenido) {
        StringBuilder consulta = new StringBuilder();
        consulta.append(" insert into t_apex_cargaarchivos (entry,ruta,archivo,extension,hoja,linea,separadores,contenido)")
                .append(" values (seq_apex_cargaarchivos.nextval,?,?,?,?,?,?,?)");

        PreparedStatement preparedStatement = null;
        try {
            preparedStatement = conexion.prepareStatement(consulta.toString());
            preparedStatement.setString(1, ruta.toLowerCase());
            preparedStatement.setString(2, archivo.toLowerCase());
            preparedStatement.setString(3, extension.toLowerCase());
            preparedStatement.setString(4, hoja.toLowerCase());
            preparedStatement.setInt(5, linea.intValue());
            preparedStatement.setInt(6, separadores.intValue());
            preparedStatement.setString(7, contenido);

            int resultado = preparedStatement.executeUpdate();
            if (resultado <= 0) {
                throw new InsertarRegistro();
            }
        } catch (Exception e) {
            StringBuilder mensajeError = new StringBuilder();
            mensajeError.append("[")
                    .append(ruta.toLowerCase())
                    .append(",")
                    .append(archivo.toLowerCase())
                    .append(",")
                    .append(linea.toString())
                    .append(",")
                    .append(contenido)
                    .append("]");
            System.out.println("Error registrar linea: " + mensajeError);

            e.printStackTrace();
        } finally {
            cerrarPreparedStatement(preparedStatement);
        }
    }

    public void insertarLogCargaArchivos(Connection conexion, String aplicacion, Integer paso, String descPaso, String observacion, String correo) {
        StringBuilder consulta = new StringBuilder();
        consulta.append(" insert into t_dlc_log(aplicacion, numpaso, despaso, observacion, correo)")
                .append(" values (?,?,?,?,?)");

        PreparedStatement preparedStatement = null;
        try {
            preparedStatement = conexion.prepareStatement(consulta.toString());
            preparedStatement.setString(1, aplicacion);
            preparedStatement.setInt(2, paso.intValue());
            preparedStatement.setString(3, descPaso);
            preparedStatement.setString(4, observacion);
            preparedStatement.setString(5, correo);

            int resultado = preparedStatement.executeUpdate();
            if (resultado <= 0) {
                throw new InsertarRegistro();
            }
        } catch (Exception e) {
            StringBuilder mensajeError = new StringBuilder();
            mensajeError.append("[")
                    .append(aplicacion)
                    .append(",")
                    .append(paso.toString())
                    .append(",")
                    .append(descPaso)
                    .append(",")
                    .append(observacion)
                    .append(",")
                    .append(correo);
            System.out.println("Error registrar linea: " + mensajeError);

            e.printStackTrace();
        } finally {
            cerrarPreparedStatement(preparedStatement);
        }
    }

    public static void cerrarResultSet(ResultSet resultSet) {
        try {
            resultSet.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void cerrarPreparedStatement(PreparedStatement preparedStatement) {
        try {
            preparedStatement.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void cerrarConexion(Connection connection) {
        try {
            connection.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
