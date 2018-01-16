package com.kbn.excel.keyword;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.robotframework.javalib.annotation.ArgumentNames;
import org.robotframework.javalib.annotation.RobotKeyword;
import org.robotframework.javalib.annotation.RobotKeywords;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;


@RobotKeywords
public class ExcelKeywords {

    XSSFWorkbook wb;
    XSSFSheet sheet;

    @RobotKeyword
    @ArgumentNames({"excelFilePath"})
    public void openExcel(String excelFilePath) {
        try {
            InputStream fileToRead = new FileInputStream(excelFilePath);
            wb = new XSSFWorkbook(fileToRead);
            sheet = wb.getSheetAt(0);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Loading file " + excelFilePath);
    }


    @RobotKeyword
    @ArgumentNames({"sheetName"})
    public void selectSheet(String sheetName) {
        sheet = wb.getSheet(sheetName);
    }

    @RobotKeyword
    @ArgumentNames({"rowNumber", "colNumber"})
    public String getCellData(int rowNumber, int colNumber) {

        XSSFRow row;
        XSSFCell cell;

        row = sheet.getRow(rowNumber);

        cell = row.getCell(colNumber);
        if (cell == null) return "";
        CellType cellType = cell.getCellTypeEnum();
        switch (cellType) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
        }

        return "";
    }

    @RobotKeyword
    @ArgumentNames({})
    public String getColumnCount() {
        return String.valueOf(sheet.getRow(0).getLastCellNum());
    }

    @RobotKeyword
    @ArgumentNames({})
    public String getRowCount() {
        return String.valueOf(sheet.getLastRowNum() + 1);
    }

    @RobotKeyword
    @ArgumentNames({"colNumber", "includeEmptyCells"})
    public String[] getColumnValues(int colNumber, boolean includeEmptyCells) {
        int rowCount = Integer.valueOf(getRowCount());
        ArrayList<String> colValues = new ArrayList<String>();

        int j = 0;
        String data = "";
        for (int i = 0; i < rowCount; i++) {
            data = getCellData(i, colNumber);
            if (!includeEmptyCells && data.equals((""))) {
                continue;
            } else {

                colValues.add(data);
            }
        }

        return colValues.toArray(new String[colValues.size()]);
    }


    @RobotKeyword
    @ArgumentNames({})
    public String getNumberOfSheets() {
        return String.valueOf(wb.getNumberOfSheets());
    }

}
