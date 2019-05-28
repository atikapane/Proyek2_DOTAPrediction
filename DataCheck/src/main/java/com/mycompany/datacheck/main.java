/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.datacheck;

import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.*;
import jxl.write.Number;

/**
 *
 * @author AtikaP
 */
public class main {
         private static final String EXCEL_FILE_LOCATION = "E:\\Kuliah\\Proyek_2\\DATAFULL.xls";
         private static final String EXCEL_WRITE_LOCATION = "E:\\Kuliah\\Proyek_2\\CHECK.xls";

    public static void main(String[] args) throws WriteException {
        int row=0, jump = 19, i=0;
        Workbook workbook = null;
        try {

            workbook = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION));
            
            Sheet sheet = workbook.getSheet(8);
            Cell cell1, cell2;
            
            WritableWorkbook myFirstWbook = null;
            myFirstWbook = Workbook.createWorkbook(new File(EXCEL_WRITE_LOCATION));
            // create an Excel sheet
            WritableSheet excelSheet = myFirstWbook.createSheet("2012", 0);
            Label label ;
            Number number;
            for(row = 0; row< 8878; row++){
                cell1 = sheet.getCell(0, row);
//                System.out.print(cell1.getContents() + "\n");   
                label = new Label(0, i, cell1.getContents().toString());
                excelSheet.addCell(label);
                number = new Number(1, i, row);
                excelSheet.addCell(number);
                i++;
                row+= jump;
            }
            myFirstWbook.write();
            if (myFirstWbook != null) {
                try {
                    myFirstWbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
            }

//            Cell cell3 = sheet.getCell(1, 0);
//            System.out.print(cell3.getContents() + ":");    // Result + :
//            Cell cell4 = sheet.getCell(1, 1);
//            System.out.println(cell4.getContents());        // Passed
//
//            System.out.print(cell1.getContents() + ":");    // Test Count + :
//            cell2 = sheet.getCell(0, 2);
//            System.out.println(cell2.getContents());        // 2
//
//            System.out.print(cell3.getContents() + ":");    // Result + :
//            cell4 = sheet.getCell(1, 2);
//            System.out.println(cell4.getContents());        // Passed 2

        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } finally {

            if (workbook != null) {
                workbook.close();
            }
        }
    }
}
