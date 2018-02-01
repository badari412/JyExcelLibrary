package com.kbn.excel.keyword;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.Arrays;
import java.util.Collection;

@RunWith(Parameterized.class)
public class ExcelKeywordsTest {

    private ExcelKeywords excelKeywords = new ExcelKeywords();

    private String originalXLSXFile = System.getProperty("user.dir").concat("\\demo\\demo.xlsx");
    private String originalXLSFile = System.getProperty("user.dir").concat("\\demo\\demo.xls");

    private static String testXLSXFile = System.getProperty("java.io.tmpdir").concat("\\test.xlsx");
    private static String testXLSFile = System.getProperty("java.io.tmpdir").concat("\\test.xls");

    @Parameterized.Parameter
    public String excelFile;

    @Parameterized.Parameters
    public static Collection<Object[]> data() {
        Object[][] data = new Object[][]{{testXLSFile}, {testXLSXFile}};
        return Arrays.asList(data);
    }


    @Before
    public void setup() {

        try {
            Files.copy(new File(originalXLSXFile).toPath(), new File(testXLSXFile).toPath(), StandardCopyOption.REPLACE_EXISTING);
            Files.copy(new File(originalXLSFile).toPath(), new File(testXLSFile).toPath(), StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    @Test
    public void testGetCellDataForNumber() throws FileNotFoundException {
        excelKeywords.openExcel(excelFile);
        excelKeywords.selectSheet("Sheet1");
        Assert.assertEquals("The value should be 1000", "1000", excelKeywords.getCellData(2, 6));

    }


    @Test
    public void testGetCellDataForString() throws FileNotFoundException {
        excelKeywords.openExcel(excelFile);
        excelKeywords.selectSheet("Sheet1");
        Assert.assertEquals("The value should be Carol", "Carol", excelKeywords.getCellData(2, 1));

    }

    @Test
    public void testGetCellDataForFormula() throws FileNotFoundException {
        excelKeywords.openExcel(excelFile);
        excelKeywords.selectSheet("Sheet1");
        Assert.assertEquals("The value should be 4600", "4600", excelKeywords.getCellData(2, 7));

    }

    @Test
    public void testGetCellDataForDecimal() throws FileNotFoundException {
        excelKeywords.openExcel(excelFile);
        excelKeywords.selectSheet("Sheet1");
        Assert.assertEquals("The value should be 40.35", "40.35", excelKeywords.getCellData(2, 4));

    }

    @Test
    public void testAddNewSheet() throws IOException {
        excelKeywords.openExcel(excelFile);
        excelKeywords.addNewSheet("TC_2");
        excelKeywords.selectSheet("TC_2");
        String[] sheetNames = excelKeywords.getSheetNames();
        boolean isCreated = false;
        for (String sheetName : sheetNames) {
            if (sheetName.equals("TC_2")) {
                isCreated = true;
                break;
            }
        }
        excelKeywords.saveExcel();
        Assert.assertTrue("A sheet should have got created with name TC_2", isCreated);

        boolean isDeleted = true;
        excelKeywords.removeSheet("TC_2");
        sheetNames = excelKeywords.getSheetNames();
        for (String sheetName : sheetNames) {
            if (sheetName.equals("TC_2")) {
                isDeleted = false;
                break;
            }
        }
        excelKeywords.saveExcel();
        Assert.assertTrue("A sheet should have got deleted with name TC_2", isDeleted);
    }


    @Test
    public void testSetCellValueWithNumber() {
        excelKeywords.openExcel(excelFile);

        excelKeywords.setCellValueWithNumber(22, 3, 9);
        excelKeywords.saveExcel();
        Assert.assertEquals("22", excelKeywords.getCellData(3, 9));

        excelKeywords.setCellValueWithNumber(2.0, 3, 9);
        excelKeywords.saveExcel();
        Assert.assertEquals("2", excelKeywords.getCellData(3, 9));

        excelKeywords.setCellValueWithNumber(22.45, 3, 9);
        excelKeywords.saveExcel();
        Assert.assertEquals("22.45", excelKeywords.getCellData(3, 9));

    }


    @Test
    public void testSetCellValueWithString() {
        excelKeywords.openExcel(excelFile);

        excelKeywords.setCellValueWithString("dummy", 3, 9);
        excelKeywords.saveExcel();
        Assert.assertEquals("dummy", excelKeywords.getCellData(3, 9));
    }


    @Test
    public void createNewWorkBook() throws InvalidFormatException {

        excelKeywords.createWorkBook(excelFile);
        excelKeywords.openExcel(excelFile);
        String[] sheets = excelKeywords.getSheetNames();
        boolean isCreated = false;
        for (String sheet : sheets) {
            if (sheet.equals("Sheet1")) {
                isCreated = true;
                break;
            }
        }

        Assert.assertTrue("Failed to create workbook " + excelFile, isCreated);
    }


    @Test
    public void testGetRowValues() {
        excelKeywords.openExcel(excelFile);
        excelKeywords.selectSheet("Sheet1");
        String[] expectedData, actualData;

        // Ignore Empty cells
        expectedData = new String[]{"3", "Rick", "35", "M", "60", "1000", "1000"};
        actualData = excelKeywords.getRowValues(3, false);
        Assert.assertTrue("Arrays are not equal.", Arrays.equals(expectedData, actualData));


        // Include Empty cells
        expectedData = new String[]{"3", "Rick", "35", "M", "60", "1000", "", "1000"};
        actualData = excelKeywords.getRowValues(3, true);
        Assert.assertTrue("Arrays are not equal.", Arrays.equals(expectedData, actualData));

    }


    @Test
    public void testGetColValues() {
        excelKeywords.openExcel(excelFile);
        excelKeywords.selectSheet("Sheet1");
        String[] expectedData, actualData;

        // Ignore empty cells
        expectedData = new String[]{"Secondary Income", "2000", "1000"};
        actualData = excelKeywords.getColumnValues(6, false);
        Assert.assertTrue("Arrays are not equal.", Arrays.equals(expectedData, actualData));

        // Include empty cells
        expectedData = new String[]{"Secondary Income", "2000", "1000", ""};
        actualData = excelKeywords.getColumnValues(6, true);
        Assert.assertTrue("Arrays are not equal.", Arrays.equals(expectedData, actualData));

    }


    @Test
    public void testSetCellWithFormula() {
        excelKeywords.openExcel(excelFile);
        excelKeywords.setCellValueWithFormula("SUM(F4,E4)", 3, 9);
        excelKeywords.saveExcel();
        Assert.assertEquals("1060", excelKeywords.getCellData(3, 9));

    }


    @Test
    public void testSetCellWithBoolean() {
        excelKeywords.openExcel(excelFile);
        excelKeywords.setCellValueWithBoolean("true", 3, 9);
        excelKeywords.saveExcel();
        Assert.assertEquals("true", excelKeywords.getCellData(3, 9));

    }

    @Test
    public void testSetCellValueWithDate() {
        excelKeywords.openExcel(excelFile);
        excelKeywords.setCellValueWithDate("12-30-2018", 3, 9);
        excelKeywords.saveExcel();
        Assert.assertEquals("12-30-2018", excelKeywords.getCellData(3, 9));
    }

    @Test
    public void testRemoveCellValue() {
        excelKeywords.openExcel(excelFile);
        excelKeywords.selectSheet("Sheet1");
        Assert.assertEquals("The value should be 40.35", "40.35", excelKeywords.getCellData(2, 4));
        excelKeywords.removeCellValue(2, 4);
        excelKeywords.saveExcel();
        Assert.assertEquals("The cell should be empty.", "", excelKeywords.getCellData(2, 4));


    }


    @After
    public void tearDown() {

        try {
            Files.delete(new File(testXLSXFile).toPath());
            Files.delete(new File(testXLSFile).toPath());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}
