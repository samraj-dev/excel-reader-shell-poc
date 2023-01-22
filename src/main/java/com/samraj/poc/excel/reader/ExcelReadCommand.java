package com.samraj.poc.excel.reader;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.shell.standard.ShellComponent;
import org.springframework.shell.standard.ShellMethod;

import java.io.*;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;

@ShellComponent
@Slf4j
public class ExcelReadCommand {

    @ShellMethod("Read Excel File")
    public String read(String filePath) {
        List<List<String>> data = null;
        String output = null;
        try {
            data = parseExcelFile(filePath);
            output = "File " + filePath + " read successfully, Data is following\r\n";
            output = data.stream().map(row -> row.toString()).collect(Collectors.joining("\r\n"));
        } catch (FileNotFoundException e) {
            log.error("File not found", e);
            output = "File not found " + e.getMessage();
        } catch (IOException e) {
            log.error("Error while processing the excel file", e);
            output = "Error while processing the excel file " + e.getMessage();
        }
        return output;
    }

    /**
     * Parse excel file and set the data in a list
     *
     * @param filePath
     * @return
     * @throws IOException
     */
    private List<List<String>> parseExcelFile(String filePath) throws IOException {
        List<List<String>> data = new ArrayList<>();
        File excel = new File(filePath);
        try (InputStream excelStream = new FileInputStream(excel)) {
            XSSFWorkbook workbook = new XSSFWorkbook(excelStream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                XSSFRow row = sheet.getRow(i);
                List<String> rowData = new ArrayList<>();
                for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                    rowData.add(row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                }
                data.add(rowData);
            }
        } catch (IOException e) {
            throw e;
        }
        return data;
    }
}
