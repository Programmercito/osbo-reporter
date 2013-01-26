/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package org.osbo.reporter;
/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

import java.sql.Connection;
import java.util.Map;

/**
 *
 * @author elJefe
 */
public class Generador {

    public void Generador() {
    }

    public boolean generar(String tipo, Map parametros, Connection conn, String reporte) {
        try {
            if (tipo.equals("PDF")) {
                ExporterPdf reportepdf = new ExporterPdf();
                reportepdf.generar(parametros, conn, reporte);
                conn.close();
                conn = null;
            }
            if (tipo.equals("XLS")) {
                ExporterXls reportepdf = new ExporterXls();
                reportepdf.generar(parametros, conn, reporte);
                conn.close();
                conn = null;
            }
            if (tipo.equals("XLSX")) {
                ExporterXlsx reportepdf = new ExporterXlsx();
                reportepdf.generar(parametros, conn, reporte);
                conn.close();
                conn = null;
            }
            if (tipo.equals("DOCX")) {
                ExporterDocx reportepdf = new ExporterDocx();
                reportepdf.generar(parametros, conn, reporte);
                conn.close();
                conn = null;
            }

        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return false;
    }
}
