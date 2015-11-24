/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mx.edu.inee.reportabe;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import javax.faces.bean.ViewScoped;
import javax.faces.context.ExternalContext;
import javax.faces.context.FacesContext;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Abel Molina S
 */
@javax.faces.bean.ManagedBean(name = "manejaArchivo")
@ViewScoped
public class ManejaArchivo implements Serializable {

    public void abrirArchivo() throws FileNotFoundException, IOException {
        FileInputStream file = new FileInputStream(new File("D:\\amolina\\Documents\\reporte_general_contraloria_interna.xlsx"));
        XSSFWorkbook wb = new XSSFWorkbook(file);
        XSSFSheet sheet = wb.getSheetAt(0);

        
        XSSFCell cell;
        int filas = 6;
        List<Peticiones> lpeticiones = llenarPeticiones();
        try {
            if (lpeticiones != null) {
                for (Peticiones p : lpeticiones) {
                    XSSFRow row = sheet.createRow(filas);
                    
                    cell = row.createCell(0);
                    cell.setCellValue(p.getA1());
                    cell = row.createCell(1);
                    cell.setCellValue(p.getA2());
                    cell = row.createCell(2);
                    cell.setCellValue(p.getA3());
                    cell = row.createCell(3);
                    cell.setCellValue(p.getA4());
                    cell = row.createCell(4);
                    cell.setCellValue(p.getA5());
                    cell = row.createCell(5);
                    cell.setCellValue(p.getA6());
                    cell = row.createCell(6);
                    cell.setCellValue(p.getA7());
                    cell = row.createCell(7);
                    cell.setCellValue(p.getA8());
                    cell = row.createCell(8);
                    cell.setCellValue(p.getA9());
                    cell = row.createCell(9);
                    cell.setCellValue(p.getA10());
                    cell = row.createCell(10);
                    cell.setCellValue(p.getA11());
                    cell = row.createCell(11);
                    cell.setCellValue(p.getA12());
                    cell = row.createCell(12);
                    cell.setCellValue(p.getA13());
                    cell = row.createCell(13);
                    cell.setCellValue(p.getA14());
                    cell = row.createCell(14);
                    cell.setCellValue(p.getA15());
                    cell = row.createCell(15);
                    cell.setCellValue(p.getA16());
                    cell = row.createCell(16);
                    cell.setCellValue(p.getA17());
                    cell = row.createCell(17);
                    cell.setCellValue(p.getA18());
                    cell = row.createCell(18);
                    cell.setCellValue(p.getA19());
                    cell = row.createCell(19);
                    cell.setCellValue(p.getA20());
                    cell = row.createCell(20);
                    cell.setCellValue(p.getA21());
                    cell = row.createCell(21);
                    cell.setCellValue(p.getA22());
                    
                    filas++;
                    
                }
                
                //write this workbook to an Outputstream.
                try (FileOutputStream fileOut = new FileOutputStream("D:\\amolina\\Documents\\reporte_general_contraloria_interna2.xlsx")) {
                    //write this workbook to an Outputstream.
                    /* Sets the password for the sheet */
                    //sheet.protectSheet("admin"); //Esta Linea se le pone contraseña al documento
                    wb.write(fileOut);
                    fileOut.flush();
                   // descargarArchivo(wb);
                }
            }
        } finally {
            file.close();
        }
    }
    
    private List<Peticiones> llenarPeticiones() {
        List<Peticiones> lpeticiones = new ArrayList<>();
        Peticiones peticiones;
        for (int i = 0; i < 10; i++) {
            peticiones = new Peticiones();
            peticiones.setA1(1);
            peticiones.setA2(2);
            peticiones.setA3(3);
            peticiones.setA4(4);
            peticiones.setA5(5);
            peticiones.setA6(6);
            peticiones.setA7(7);
            peticiones.setA8(8);
            peticiones.setA9(9);
            peticiones.setA10(10);
            peticiones.setA11(11);
            peticiones.setA12(12);
            peticiones.setA13(13);
            peticiones.setA14(14);
            peticiones.setA15(15);
            peticiones.setA16(16);
            peticiones.setA17(17);
            peticiones.setA18(18);
            peticiones.setA19(19);
            peticiones.setA20(20);
            peticiones.setA21(21);
            peticiones.setA22(22);
            lpeticiones.add(peticiones);
            
        }
        
        return lpeticiones;
    }
    
    /**
     * Este metodo descarga el archivo en el navegador
     * @param wb
     * @throws IOException 
     */
    private void descargarArchivo(XSSFWorkbook wb) throws IOException {
        FacesContext context = FacesContext.getCurrentInstance();
        ExternalContext externalContext = context.getExternalContext();
        externalContext.responseReset();        
        externalContext.setResponseContentType("application/vnd.ms-excel");
        externalContext.setResponseHeader("Content-Disposition", "attachment;filename=export.xlsx");
        wb.write(externalContext.getResponseOutputStream());
        context.responseComplete();        
    }
}
