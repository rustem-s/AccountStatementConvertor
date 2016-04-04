package ru.mobilewin.accountstatementconvertor;

import com.aspose.cells.Workbook;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

/**
 * Created by Rustem.Saidaliyev on 01.04.2016.
 */
public class Utils {
    public static String convertFile(String file) throws Exception {

        String amountCurrent = "";
        String stringNew = "";
        boolean deleteFile = false;

        String fileExtension = FilenameUtils.getExtension(file);
        if (fileExtension.equalsIgnoreCase("xlsb")) {
            Workbook workbook = new Workbook(file);
            String fileWithoutExt = FilenameUtils.removeExtension(file);
            file = fileWithoutExt + ".xlsx";
            deleteFile = true;
            workbook.save(file);
        }
        FileInputStream fileInputStream = new FileInputStream(file);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);

        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);

        Iterator<Row> rowIterator = xssfSheet.rowIterator();

        int rowCount = 0;

        while (rowIterator.hasNext()) {
            XSSFRow xssfRow = (XSSFRow) rowIterator.next();
            if (xssfRow.getRowNum() == 0) {
                continue;
            }
            if (xssfRow.getRowNum() > 11 && xssfRow.getRowNum() < xssfSheet.getLastRowNum() && xssfRow.getCell(9).toString() != "") {
                amountCurrent = xssfRow.getCell(9).getRawValue().toString().replace("\t", " ").trim();
                if (!amountCurrent.contains(".")) {
                    amountCurrent = amountCurrent + ".00";
                }
                stringNew = stringNew
                        + xssfRow.getCell(1).toString().replace("\t", " ").trim() + "\t"
                        + xssfRow.getCell(2).toString().replace("\t", " ").trim() + "\t"
                        + xssfRow.getCell(5).getRawValue().toString().replace("\t", " ").trim() + "\t"
                        + xssfRow.getCell(6).toString().replace("\t", " ").trim() + "\t"
                        + xssfRow.getCell(7).toString().replace("\t", " ").trim() + "\t"
                        + amountCurrent + "\t"
                        + xssfRow.getCell(4).toString().replace("\t", " ").trim() + "\t"
                        + xssfRow.getCell(10).toString().replace("\t", " ").trim() + "\n";

            }
            rowCount++;
        }
        String header = "Date-dd.mm.yyyy,Nr,INN,BIK,Account,Amount,Payer,Purpose of payment," + (rowCount + 2) + "\n";

        if (deleteFile) {
            File f = new File(file);
            f.delete();
        }

        return header + stringNew;
    }

    /*
    public static String convertFile2(String file) throws Exception {

        String amountCurrent = "";
        String stringNew = "";

        Workbook workbook = new Workbook(file);
        Worksheet worksheet = workbook.getWorksheets().get(0);

        RowCollection rowCollection = worksheet.getCells().getRows();

        Iterator<Row> rowIterator = rowCollection.iterator();

        Iterator<Row> rowIterator = xssfSheet.rowIterator();

        int rowCount = 0;

        while (rowIterator.hasNext()) {
            XSSFRow xssfRow = (XSSFRow) rowIterator.next();
            if (xssfRow.getRowNum() == 0) {
                continue;
            }
            if (xssfRow.getRowNum() > 11 && xssfRow.getRowNum() < worksheet.getCells().getRowgetLastRowNum() && xssfRow.getCell(9).toString() != "") {
                amountCurrent = xssfRow.getCell(9).getRawValue().toString().replace("\t", " ").trim();
                if (!amountCurrent.contains(".")) {
                    amountCurrent = amountCurrent + ".00";
                }
                stringNew = stringNew
                        + xssfRow.getCell(1).toString().replace("\t", " ").trim() + "\t"
                        + xssfRow.getCell(2).toString().replace("\t", " ").trim() + "\t"
                        + xssfRow.getCell(5).getRawValue().toString().replace("\t", " ").trim() + "\t"
                        + xssfRow.getCell(6).toString().replace("\t", " ").trim() + "\t"
                        + xssfRow.getCell(7).toString().replace("\t", " ").trim() + "\t"
                        + amountCurrent + "\t"
                        + xssfRow.getCell(4).toString().replace("\t", " ").trim() + "\t"
                        + xssfRow.getCell(10).toString().replace("\t", " ").trim() + "\n";

            }
            rowCount++;
        }
        String header = "Date-dd.mm.yyyy,Nr,INN,BIK,Account,Amount,Payer,Purpose of payment," + (rowCount + 2) + "\n";

        return header + stringNew;
    }
    */

    public static int okcancel(String theMessage) {
        int result = JOptionPane.showConfirmDialog(null, theMessage,
                "alert", JOptionPane.OK_CANCEL_OPTION);
        return result;
    }
}
