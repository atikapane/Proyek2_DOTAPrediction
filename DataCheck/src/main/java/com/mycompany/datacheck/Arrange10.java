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
 * @author AtikaP and edited by FadlyTr
 */
public class Arrange10 {

    private static final String EXCEL_FILE_LOCATION = "E:\\Kuliah\\Proyek_2\\DATAFULL.xls";
    private static final String EXCEL_WRITE_LOCATION = "E:\\Kuliah\\Proyek_2\\TEMP.xls";

    public static void main(String[] args) throws WriteException {
        int row = 0, jumpMatch = 19, rowWrite = 0, colWrite, rowT1 = 6, rowT2 = 13, i, j, k, TeamT1 = 4, TeamT2 = 11, winner = 2;
        Workbook workbook = null;
        try {

            workbook = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION));
            Sheet sheet = workbook.getSheet(8);
            Cell cell1, T1, T2, Winner;

            WritableWorkbook myFirstWbook = null;
            myFirstWbook = Workbook.createWorkbook(new File(EXCEL_WRITE_LOCATION));
            // create an Excel sheet
            WritableSheet excelSheet = myFirstWbook.createSheet("2017", 0);
            Label label;
            for (i = 0; i < 10; i++) {
                label = new Label(i, rowWrite, "Hero" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 10; i < 20; i++) {
                label = new Label(i, rowWrite, "Nama" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 20; i < 30; i++) {
                label = new Label(i, rowWrite, "K" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 30; i < 40; i++) {
                label = new Label(i, rowWrite, "D" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 40; i < 50; i++) {
                label = new Label(i, rowWrite, "A" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 50; i < 60; i++) {
                label = new Label(i, rowWrite, "NET" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 60; i < 70; i++) {
                label = new Label(i, rowWrite, "LH" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 70; i < 80; i++) {
                label = new Label(i, rowWrite, "DN" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 80; i < 90; i++) {
                label = new Label(i, rowWrite, "GPM" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 90; i < 100; i++) {
                label = new Label(i, rowWrite, "XPM" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 100; i < 110; i++) {
                label = new Label(i, rowWrite, "DMG" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 110; i < 120; i++) {
                label = new Label(i, rowWrite, "HEAL" + i % 10);
                excelSheet.addCell(label);
            }
            for (i = 120; i < 130; i++) {
                label = new Label(i, rowWrite, "BLD" + i % 10);
                excelSheet.addCell(label);
            }

            rowWrite++;
            String str = new String();
            for (row = 0; row < 8778; row+=18) {
                colWrite = 0;
                for (k = 0; k < 13; k++) {
                    
                    for (j = 0; j < 5; j++) {
                        cell1 = sheet.getCell(k, row + rowT1 + j);
                        str = cell1.getContents();
                        if (str.equals("-")) {
                            str = "0";
                        } else if (str.charAt(str.length() - 1) == 'k') {
                            str = str.substring(0, str.length() - 3) + str.substring(str.length() - 2, str.length() - 1) + "00";
                        }
                        label = new Label(colWrite, rowWrite, str);
                        excelSheet.addCell(label);
                        colWrite++;
                    }
                    
                    for (j = 0; j < 5; j++) {
                        cell1 = sheet.getCell(k, row + rowT2 + j);
                        str = cell1.getContents();
                        if (str.equals("-")) {
                            str = "0";
                        } else if (str.charAt(str.length() - 1) == 'k') {
                            str = str.substring(0, str.length() - 3) + str.substring(str.length() - 2, str.length() - 1) + "00";
                        }
                        label = new Label(colWrite, rowWrite, str);
                        excelSheet.addCell(label);
                        colWrite++;
                    }
                    
                    //Check winner
                    T1 = sheet.getCell(0, row + TeamT1);
                    T2 = sheet.getCell(0, row + TeamT2);
                    Winner = sheet.getCell(0, row + winner);
                    if (Winner.getContents().substring(0, Winner.getContents().length() - 9).equals(T1.getContents())) {
                        label = new Label(colWrite, rowWrite, "1");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "1");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "1");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "1");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "1");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "0");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "0");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "0");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "0");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "0");
                        excelSheet.addCell(label);
                        colWrite++;
                    } else {
                        label = new Label(colWrite, rowWrite, "0");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "0");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "0");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "0");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "0");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "1");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "1");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "1");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "1");
                        excelSheet.addCell(label);
                        colWrite++;
                        label = new Label(colWrite, rowWrite, "1");
                        excelSheet.addCell(label);
                        colWrite++;
                    }
                }
                rowWrite++;

                row += jumpMatch;
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
