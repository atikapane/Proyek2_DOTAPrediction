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

/**
 *
 * @author AtikaP
 */
public class Arrange {
    private static final String EXCEL_FILE_LOCATION = "E:\\Kuliah\\Proyek_2\\DATAFULL.xls";
         private static final String EXCEL_WRITE_LOCATION = "E:\\Kuliah\\Proyek_2\\TEMP.xls";

    public static void main(String[] args) throws WriteException {
        int row=0, jumpMatch = 19, rowWrite = 0, colWrite, rowT1 = 6, rowT2 = 13, i, j, k, TeamT1=4, TeamT2=11, winner=2;
        Workbook workbook = null;
        try {

            workbook = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION));
            Sheet sheet = workbook.getSheet(8);
            Cell cell1, T1, T2, Winner;
            
            
            WritableWorkbook myFirstWbook = null;
            myFirstWbook = Workbook.createWorkbook(new File(EXCEL_WRITE_LOCATION));
            // create an Excel sheet
            WritableSheet excelSheet = myFirstWbook.createSheet("2017", 0);
            Label label ;
            label = new Label(0, rowWrite, "name0");
            excelSheet.addCell(label);
            label = new Label(1, rowWrite, "name1");
            excelSheet.addCell(label);
            label = new Label(2, rowWrite, "K0");
            excelSheet.addCell(label);
            label = new Label(3, rowWrite, "K1");
            excelSheet.addCell(label);
            label = new Label(4, rowWrite, "D0");
            excelSheet.addCell(label);
            label = new Label(5, rowWrite, "D1");
            excelSheet.addCell(label);
            label = new Label(6, rowWrite, "A0");
            excelSheet.addCell(label);
            label = new Label(7, rowWrite, "A1");
            excelSheet.addCell(label);
            label = new Label(8, rowWrite, "NET0");
            excelSheet.addCell(label);
            label = new Label(9, rowWrite, "NET1");
            excelSheet.addCell(label);
            label = new Label(10, rowWrite, "LH0");
            excelSheet.addCell(label);
            label = new Label(11, rowWrite, "LH1");
            excelSheet.addCell(label);
            label = new Label(12, rowWrite, "DN0");
            excelSheet.addCell(label);
            label = new Label(13, rowWrite, "DN1");
            excelSheet.addCell(label);
            label = new Label(14, rowWrite, "GPM0");
            excelSheet.addCell(label);
            label = new Label(15, rowWrite, "GPM1");
            excelSheet.addCell(label);
            label = new Label(16, rowWrite, "XPM0");
            excelSheet.addCell(label);
            label = new Label(17, rowWrite, "XPM1");
            excelSheet.addCell(label);
            label = new Label(18, rowWrite, "DMG0");
            excelSheet.addCell(label);
            label = new Label(19, rowWrite, "DMG1");
            excelSheet.addCell(label);
            label = new Label(20, rowWrite, "HEAL0");
            excelSheet.addCell(label);
            label = new Label(21, rowWrite, "HEAL1");
            excelSheet.addCell(label);
            label = new Label(22, rowWrite, "BLD0");
            excelSheet.addCell(label);
            label = new Label(23, rowWrite, "BLD1");
            excelSheet.addCell(label);
            label = new Label(24, rowWrite, "WIN0");
            excelSheet.addCell(label);
            label = new Label(25, rowWrite, "WIN1");
            excelSheet.addCell(label);
            rowWrite++;
            for(row = 0; row< 8778; row++){
                for(i = 0; i<5; i++){
                    for(j = 0; j<5; j++){
                        colWrite = 0;
                        for(k = 0; k<12; k++){
                            cell1 = sheet.getCell(k+1, row+rowT1+i);
                            label = new Label(colWrite, rowWrite, cell1.getContents());
                            excelSheet.addCell(label);
                            colWrite++;
                            cell1 = sheet.getCell(k+1, row+rowT2+j);
                            label = new Label(colWrite, rowWrite, cell1.getContents());
                            excelSheet.addCell(label);
                            colWrite++;
                        }
                        //Check winner
                        T1 = sheet.getCell(0, row+TeamT1);
                        T2 = sheet.getCell(0, row+TeamT2);
                        Winner = sheet.getCell(0, row+winner);
                        if(Winner.getContents().substring(0, Winner.getContents().length()-9).equals(T1.getContents())){
                            label = new Label(colWrite, rowWrite,"1");
                            excelSheet.addCell(label);
                            colWrite++;
                            label = new Label(colWrite, rowWrite,"0");
                            excelSheet.addCell(label);
                            colWrite++;
                        } else {
                            label = new Label(colWrite, rowWrite,"0");
                            excelSheet.addCell(label);
                            colWrite++;
                            label = new Label(colWrite, rowWrite,"1");
                            excelSheet.addCell(label);
                            colWrite++;
                        }
                        rowWrite++;
                    }
                }
                row+= jumpMatch;
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
