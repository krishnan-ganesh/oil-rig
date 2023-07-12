package com.kg;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

public class XLSXToJson {
    public static List<Map<String,String>> asListOfMaps(String xlsxFilePath,String workSheet){
        DataFormatter formatter = new DataFormatter();
        HashMap<Integer,String> headers = new HashMap<>();
        List<Map<String,String>> testCases = new LinkedList<>();

        try(FileInputStream fileInputStream = new FileInputStream(xlsxFilePath)){
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
            boolean rowIsHeader = true;
            boolean headerFound = false;
            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(xssfWorkbook.getSheetIndex(workSheet));
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
            return testCases;

        }
        catch (IOException e){
            e.printStackTrace();
            return null;
        }
    }

    public static String asJsonString(String xlsxFilePath,String workSheet){
        try {
            return new ObjectMapper().writeValueAsString(asListOfMaps(xlsxFilePath,workSheet));
        } catch (JsonProcessingException e) {
            e.printStackTrace();
            return "{}";
        }
    }
}
