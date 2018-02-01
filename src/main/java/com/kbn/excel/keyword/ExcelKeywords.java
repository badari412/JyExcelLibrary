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


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.robotframework.javalib.annotation.ArgumentNames;
import org.robotframework.javalib.annotation.RobotKeyword;
import org.robotframework.javalib.annotation.RobotKeywords;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.regex.Pattern;


@RobotKeywords
public class ExcelKeywords {

    private InputStream fileInputStream;
    private OutputStream fileOutputStream;
    private Workbook wb;
    private Sheet sheet;
    private String excelFilePath;


    @RobotKeyword("Open the excel file using the given path.\n\n" +
            "Example:\n" +
            "| Open Excel | C:\\\\demo.xlsx |" +
            "\n")
    @ArgumentNames({"excelFilePath"})
    public void openExcel(String excelFilePath) {
        this.excelFilePath = excelFilePath.trim();
        openFileToRead();
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

        Row row;
        Cell cell;
        DataFormatter dataFormatter = new DataFormatter();
        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
        row = sheet.getRow(rowNumber);

        cell = row.getCell(colNumber);
        if (cell == null) {
            return "";
        } else {
            return dataFormatter.formatCellValue(cell, formulaEvaluator);
        }

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
        ArrayList<String> colValues = new ArrayList<>();

        String data;
        for (int i = 0; i < rowCount; i++) {
            data = getCellData(i, colNumber);
            if (!(!includeEmptyCells && data.equals(""))) {
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

        ArrayList<String> rowValues = new ArrayList<>();
        int colCount = Integer.parseInt(getColumnCount());
        String data;

        for (int i = 0; i < colCount; i++) {
            data = getCellData(rowNumber, i);

            if (!includeEmptyCells && data.equals("")) {
                continue;
            }
            rowValues.add(data);


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
        ArrayList<String> sheetNames = new ArrayList<>();
        for (int i = 0; i < noOfSheets; i++) {
            sheetNames.add(wb.getSheetName(i));
        }

        return sheetNames.toArray(new String[sheetNames.size()]);
    }

    @RobotKeyword("Add a new sheet to the currently opened excel.\n\n" +
            "Example:\n" +
            "| Add New Sheet | Sheet2 |\n" +
            "\n")
    @ArgumentNames({"name"})
    public void addNewSheet(String name) throws IOException {

        wb.createSheet(name);


    }

    @RobotKeyword("Sets the value of the cell in the active sheet with a number.\n\n" +
            "Example:\n" +
            "| Set Cell Value With Number | 34 | 1 | 2 |\n" +
            "| Set Cell Value With Number | 34.3 | 1 | 2 |\n" +
            "| Set Cell Value With Number | 34.59 | 1 | 2 |\n" +
            "\n")
    @ArgumentNames({"number", "rowNumber", "columnNumber"})
    public void setCellValueWithNumber(double number, int rowNumber, int columnNumber) {

        Cell cell = getCell(rowNumber, columnNumber);
        cell.setCellType(CellType.NUMERIC);
        cell.setCellValue(number);

    }


    @RobotKeyword("Sets the value of the cell in the active sheet with a string.\n\n" +
            "Example:\n" +
            "| Set Cell Value With String | dummy | 1 | 2 |\n" +
            "\n")
    @ArgumentNames({"string", "rowNumber", "columnNumber"})
    public void setCellValueWithString(String string, int rowNumber, int columnNumber) {

        Cell cell = getCell(rowNumber, columnNumber);
        cell.setCellType(CellType.STRING);
        cell.setCellValue(string);

    }


    @RobotKeyword("Sets the value of the cell in the active sheet with a formula.\n\n" +
            "Example:\n" +
            "| Set Cell Value With Formula | SUM(F1,G1) | 1 | 2 |\n" +
            "\n")
    @ArgumentNames({"string", "rowNumber", "columnNumber"})

    public void setCellValueWithFormula(String formula, int rowNumber, int columnNumber) {
        Cell cell = getCell(rowNumber, columnNumber);
        cell.setCellType(CellType.FORMULA);
        cell.setCellFormula(formula);

    }

    @RobotKeyword("Removes the given sheet from the active workbook.\n\n" +
            "Example:\n" +
            "| Remove Sheet | TC_2 |\n" +
            "\n")
    @ArgumentNames({"string"})
    public void removeSheet(String sheetName) {
        wb.removeSheetAt(wb.getSheetIndex(sheetName));

    }


    @RobotKeyword("Sets the value of the cell in the active sheet with a boolean value.\n\n" +
            "Example:\n" +
            "| Set Cell Value With Boolean | true | 1 | 2 |\n" +
            "\n")
    @ArgumentNames({"string", "rowNumber", "columnNumber"})
    public void setCellValueWithBoolean(String booleanValue, int rowNumber, int columnNumber) {
        Cell cell = getCell(rowNumber, columnNumber);
        cell.setCellType(CellType.BOOLEAN);
        cell.setCellValue(booleanValue);
    }


    @RobotKeyword("Sets the value of the cell in the active sheet with date(MM-dd-yyyy).\n\n" +
            "Example:\n" +
            "| Set Cell Value With Date | 03-30-2018 | 1 | 2 |\n" +
            "\n")
    @ArgumentNames({"dateValue", "rowNumber", "columnNumber"})
    public void setCellValueWithDate(String dateValue, int rowNumber, int columnNumber) {

        CellStyle cellStyle = wb.createCellStyle();
        CreationHelper createHelper = wb.getCreationHelper();
        short dateFormat = createHelper.createDataFormat().getFormat("MM-dd-yyyy");
        cellStyle.setDataFormat(dateFormat);

        SimpleDateFormat formatter = new SimpleDateFormat("MM-dd-yyyy");

        Cell cell = getCell(rowNumber, columnNumber);
        try {
            cell.setCellValue(formatter.parse(dateValue));
        } catch (ParseException e) {
            e.printStackTrace();
        }
        cell.setCellStyle(cellStyle);

    }


    @RobotKeyword("Removes the value from the given cell in the active sheet.\n\n" +
            "Example:\n" +
            "| Remove Cell Value | 1 | 2 |\n" +
            "\n")
    @ArgumentNames({"rowNumber", "columnNumber"})
    public void removeCellValue(int rowNumber, int columnNumber) {
        Cell cell = getCell(rowNumber, columnNumber);
        cell.setCellType(CellType.BLANK);
        cell.setCellValue("");


    }

    @RobotKeyword("Saves the excel sheet after making any changes to it.\n\n" +
            "Example:\n" +
            "| Remove Cell Value | 1 | 2 |\n" +
            "| Save Excel |\n")
    @ArgumentNames({})

    public void saveExcel() {
        openFileToWrite();
    }

    private Cell getCell(int rowNumber, int columnNumber) {
        Row row = sheet.getRow(rowNumber);
        Cell cell;
        if (row == null) {
            row = sheet.createRow(rowNumber);
        }

        cell = row.getCell(columnNumber);
        if (cell == null) {
            cell = row.createCell(columnNumber);
        }
        return cell;
    }

    @RobotKeyword("Creates a new excel workbook with the given name.\n\n" +
            "Example:\n" +
            "| Create Workbook | C:\\Demo.xlsx |\n" +
            "| Create Workbook | C:\\Demo.xls |\n" +
            "\n")
    @ArgumentNames({"rowNumber", "columnNumber"})
    public void createWorkBook(String excelFilePath) throws InvalidFormatException {
        this.excelFilePath = excelFilePath;
        Workbook newWb;

        if (excelFilePath.endsWith(".xlsx")) {
            newWb = new XSSFWorkbook();
        } else if (excelFilePath.endsWith(".xls")) {
            newWb = new HSSFWorkbook();
        } else {
            throw new InvalidFormatException("Please make sure you use either .xlsx or .xls format.");
        }


        newWb.createSheet("Sheet1");

        try {
            FileOutputStream newFOStream = new FileOutputStream(excelFilePath);
            newWb.write(newFOStream);
            newFOStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private void openFileToRead() {

        try {
            fileInputStream = new FileInputStream(excelFilePath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        try {
            wb = WorkbookFactory.create(fileInputStream);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
        sheet = wb.getSheetAt(0);
    }

    private void openFileToWrite() {

        try {
            fileOutputStream = new FileOutputStream(excelFilePath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            wb.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            fileOutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
