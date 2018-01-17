/*
 * Copyright 2018 Badari Narayana
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

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


    @RobotKeyword("Open the excel file using the given path.\n\n" +
            "Example:\n" +
            "| Open Excel | C:\\\\demo.xlsx |" +
            "\n")
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


    @RobotKeyword("Selects the sheet by its name.\n\n" +
            "Example:\n" +
            "| Select Sheet | Demo |" +
            "\n")
    @ArgumentNames({"sheetName"})
    public void selectSheet(String sheetName) {
        sheet = wb.getSheet(sheetName);
    }


    @RobotKeyword("Gets the data from the active sheet.\n" +
            "Uses the cell's row number and column number.\n" +
            "Example:\n" +
            "| Get Cell Data  | 1 | 2 |" +
            "\n")
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


    @RobotKeyword("Returns the number of columns from the active sheet.\n\n" +
            "Example:\n" +
            "| ${colCount} | Get Column Count |\n" +
            "| Should Be Equal As Integers | 2 | ${colCount} |" +
            "\n")
    @ArgumentNames({})
    public String getColumnCount() {
        return String.valueOf(sheet.getRow(0).getLastCellNum());
    }


    @RobotKeyword("Returns the number of rows from the active sheet.\n\n" +
            "Example:\n" +
            "| ${rowCount} | Get Row Count |\n" +
            "| Should Be Equal As Integers | 2 | ${rowCount} |" +
            "\n")
    @ArgumentNames({})
    public String getRowCount() {
        return String.valueOf(sheet.getLastRowNum() + 1);
    }


    @RobotKeyword("Returns the column values of a given column (using its index) from the active sheet.\n\n" +
            "Example:\n" +
            "| ${result} | Get Column Values | 2 | True |\n" +
            "| Should Be Equal As Strings | Demo | ${result[0]} |" +
            "\n")
    @ArgumentNames({"colNumber", "includeEmptyCells"})
    public String[] getColumnValues(int colNumber, boolean includeEmptyCells) {
        int rowCount = Integer.valueOf(getRowCount());
        ArrayList<String> colValues = new ArrayList<String>();

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


    @RobotKeyword("Returns the number of sheets present in the currently opened excel file.\n\n" +
            "Example:\n" +
            "| ${result} | Get Number Of Sheets |\n" +
            "| Should Be Equal As Integers | 2 | ${result} |" +
            "\n")
    @ArgumentNames({})
    public String getNumberOfSheets() {
        return String.valueOf(wb.getNumberOfSheets());
    }


    @RobotKeyword("Returns the row values of a given row (using it's index) from the active sheet.\n\n" +
            "Example:\n" +
            "| ${result} | Get Row Values | 2 | True |\n" +
            "| Should Be Equal As Strings | Demo | ${result[0]} |" +
            "\n")
    @ArgumentNames({"rowNumber", "includeEmptyCells"})
    public String[] getRowValues(int rowNumber, boolean includeEmptyCells) {

        XSSFRow row = sheet.getRow(rowNumber);
        ArrayList<String> rowValues = new ArrayList<String>();
        int colCount = Integer.parseInt(getColumnCount());
        String data = "";

        for (int i = 0; i < colCount; i++) {
            data = getCellData(rowNumber, i);
            if (!includeEmptyCells && data.equals("")) {
                continue;
            } else {
                rowValues.add(data);
            }

        }

        return rowValues.toArray(new String[rowValues.size()]);

    }


    @RobotKeyword("Returns a list of names of the sheets present in the currently opened excel file.\n\n" +
            "Example:\n" +
            "| ${result} | Get Sheet Names |\n" +
            "| Should Be Equal As Strings | TC_1 | ${result[1]} |" +
            "\n")
    @ArgumentNames({})
    public String[] getSheetNames() {

        int noOfSheets = Integer.parseInt(getNumberOfSheets());
        ArrayList<String> sheetNames = new ArrayList<String>();
        for (int i = 0; i < noOfSheets; i++) {
            sheetNames.add(wb.getSheetName(i));
        }

        return sheetNames.toArray(new String[sheetNames.size()]);
    }

}
