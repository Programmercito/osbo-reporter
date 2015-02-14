/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package org.osbo.reporter;

import java.io.InputStream;
import java.sql.Connection;
import java.util.Map;
import javax.faces.context.ExternalContext;
import javax.faces.context.FacesContext;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import net.sf.jasperreports.engine.JRParameter;
import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;

import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.util.JRLoader;

/**
 *
 * @author osbosoftware creador
 */
public class ExporterPdf {

    FacesContext fc;
    ExternalContext ec;

    public void ExporterPdf() {
    }

    public void generar(Map parametros, Connection conn, String reporte) {
        try {
            fc = FacesContext.getCurrentInstance();
            ec = fc.getExternalContext();
            InputStream fis = ec.getResourceAsStream(reporte);
            JasperReport jasperReport=(JasperReport)JRLoader.loadObject(fis) ;
            parametros.put(JRParameter.IS_IGNORE_PAGINATION, Boolean.FALSE);
            JasperPrint jasperPrint = JasperFillManager.fillReport(
                   jasperReport, parametros, conn);
            HttpServletResponse resp = (HttpServletResponse) ec.getResponse();
            resp.setContentType("application/pdf");
            String filename = "resultado.pdf";
            resp.addHeader("Content-Disposition", "attachment; filename=\"" + filename+"\"");
            ServletOutputStream sou=resp.getOutputStream();
            sou.write(JasperExportManager.exportReportToPdf(jasperPrint));

            System.out.println(" reporte generado ");
            fc.getApplication().getStateManager().saveView(fc);
            fc.responseComplete();
        } catch (Exception ex) {
            System.out.println("error" + ex.getMessage());
            ex.printStackTrace();
        }

    }
}
