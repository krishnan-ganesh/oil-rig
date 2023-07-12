package com.kg;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class Rig {
    public static void main(String[] args) {
        DataFormatter formatter = new DataFormatter();
        HashMap<Integer,String> headers = new HashMap<>();
        List<Map<String,String>> testCases = new LinkedList<>();

        try(FileInputStream fileInputStream = new FileInputStream("testsheet.xlsx")){
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
            boolean rowIsHeader = true;
            boolean headerFound = false;
            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(xssfWorkbook.getSheetIndex("Test Results"));
            for (Row row : xssfSheet) {
                HashMap<String,String> value = new HashMap<>();
                for (Cell cell : row) {
                    if (cell.getCellType() != CellType._NONE && cell.getCellType() != CellType.BLANK) {
                        if (rowIsHeader) {
                            headerFound = true;
                            headers.put(cell.getColumnIndex(),formatter.formatCellValue(cell));
                        }
                        else {
                            value.put(headers.get(cell.getColumnIndex()),formatter.formatCellValue(cell));
                        }
                    }
                }
                if(rowIsHeader && headerFound){
                    rowIsHeader = false;
                }
                else{
                    if(!value.isEmpty()) testCases.add(value);
                }
            }
            System.out.println(new ObjectMapper().writeValueAsString(testCases));

        }
        catch (IOException e){
            e.printStackTrace();
        }

    }
}