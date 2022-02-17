package com.finskayaylochka.pdfreader.service;

import com.finskayaylochka.pdfreader.model.ExcelDocument;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;

/**
 * @author Alexandr Stegnin
 */
@Service
@Slf4j
public class ExcelService {

  @SneakyThrows
  public void buildExcel(String excelPath, ExcelDocument excelDocument) {
    try (InputStream inputStream = new FileInputStream(excelPath);
         XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
         OutputStream outputStream = new FileOutputStream(excelPath)) {
      Sheet sheet = workbook.getSheetAt(0);
      createRow(sheet, excelDocument);
      workbook.write(outputStream);
    } catch (IOException e) {
      log.error("Ошибка заполнения excel файла {}: {}", excelPath, e);
    }
  }

  private void createRow(Sheet sheet, ExcelDocument excelDocument) {
    int lastRow = sheet.getPhysicalNumberOfRows();
    Row row = sheet.createRow(lastRow);
    createCell(row, 0, excelDocument.getAddress());
    createCell(row, 1, excelDocument.getArea());
    createCell(row, 2, excelDocument.getSum());
    createCell(row, 3, excelDocument.getNumbers());
    createCell(row, 4, excelDocument.getNumber());
    createCell(row, 5, excelDocument.getOwner());
    createCell(row, 6, excelDocument.getInfo());
  }

  private void createCell(Row row, int index, String value) {
    Cell cell = row.createCell(index);
    cell.setCellValue(value);
  }

}
