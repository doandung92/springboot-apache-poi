package com.apache.poi.service;

import com.apache.poi.model.Data;
import com.fasterxml.jackson.annotation.JsonProperty;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.List;

@Service
public class PoiService {

    public void readData(String file) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        int rowNum = sheet.getLastRowNum();
        int colNum = sheet.getRow(0).getLastCellNum();
        for (int i = 1; i <= rowNum ; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int j = 0; j <colNum ; j++) {
                XSSFCell cell = row.getCell(j);
                switch (cell.getCellType()){
                    case STRING:
                        System.out.println("String :"+ cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell))
                            System.out.println("Number :"+ cell.getDateCellValue());
                        else
                        System.out.println("Number :"+ cell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        System.out.println("Boolean :"+ cell.getBooleanCellValue());
                        break;
                        // Formula for numeric
                    case FORMULA:
                        System.out.println("Formula :"+ cell.getNumericCellValue());
                        break;
                    default:
                        System.out.println("Blank");
                }
            }
        }
    }
    public void writeData(String file, List<Data> dataList) throws IOException, IllegalAccessException {
        int colNum = 3;
        FileInputStream fileInputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        int addRow = 1;
        XSSFRow headerRow = sheet.getRow(0);
        Field[] fields = Data.class.getDeclaredFields();
        for (int i = 0; i < dataList.size(); i++) {
            XSSFRow row = sheet.createRow(addRow++);
            XSSFRow row1 = sheet.createRow(addRow++);
            for (int j = 0; j < fields.length; j++) {
                fields[j].setAccessible(true);
                if (!fields[j].isAnnotationPresent(JsonProperty.class)) continue;
                row1.createCell(j).setCellValue(fields[j].getDeclaredAnnotation(JsonProperty.class).value());
                Object value = fields[j].get(dataList.get(i));
                XSSFCell cell = row.createCell(j);
                if (value instanceof Number) cell.setCellValue(Double.parseDouble(value.toString()));
                else if (value instanceof Boolean) cell.setCellValue(Boolean.parseBoolean(value.toString()));
                else cell.setCellValue(value.toString());
                cell.setCellStyle(headerRow.getCell(j).getCellStyle());
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);
        workbook.close();
    }

}
