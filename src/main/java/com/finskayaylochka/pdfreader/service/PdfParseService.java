package com.finskayaylochka.pdfreader.service;

import com.finskayaylochka.pdfreader.model.ExcelDocument;
import lombok.AccessLevel;
import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;
import lombok.experimental.FieldDefaults;
import lombok.extern.slf4j.Slf4j;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.springframework.stereotype.Service;
import technology.tabula.ObjectExtractor;
import technology.tabula.Page;
import technology.tabula.RectangularTextContainer;
import technology.tabula.Table;
import technology.tabula.extractors.SpreadsheetExtractionAlgorithm;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author Alexandr Stegnin
 */
@Service
@RequiredArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
@SuppressWarnings("all")
@Slf4j
public class PdfParseService {

  ExcelService excelService;

  public void parseDir(String dir, String excelPath) {
    collectExcelData(getFiles(dir))
        .stream()
        .filter(Objects::nonNull)
        .forEach(excelDocument -> excelService.buildExcel(excelPath, excelDocument));
  }

  @SneakyThrows
  private Set<String> getFiles(String dir) {
    try (Stream<Path> stream = Files.list(Paths.get(dir))) {
      return stream
          .filter(file -> !Files.isDirectory(file))
          .map(path -> path.toFile().getAbsolutePath())
          .collect(Collectors.toSet());
    }
  }

  private List<ExcelDocument> collectExcelData(Set<String> filenames) {
    List<ExcelDocument> documents = filenames.stream()
        .map(file -> parsePdf(file))
        .collect(Collectors.toList());
    return documents;
  }

  public ExcelDocument parsePdf(String path) {
    try (PDDocument pd = PDDocument.load(new File(path));) {
      int totalPages = pd.getNumberOfPages();
      log.info("Found pages count: {}", totalPages);

      ObjectExtractor oe = new ObjectExtractor(pd);
      SpreadsheetExtractionAlgorithm sea = new SpreadsheetExtractionAlgorithm();
      ExcelDocument excelDocument = new ExcelDocument();
      for (int i = 1; i <= totalPages; i++) {
        Page page = oe.extract(i);
        List<Table> tables = sea.extract(page);
        tables.forEach(table -> {
          var rows = table.getRows();
          for (List<RectangularTextContainer> cells : rows) {
            for (int k = 0; k < cells.size(); k++) {
              if (cells.get(k).getText().equalsIgnoreCase("кадастровый номер:")) {
                excelDocument.setNumber(cells.get(k + 1).getText());
              } else if (cells.get(k).getText().equalsIgnoreCase("адрес:")) {
                excelDocument.setAddress(cells.get(k + 1).getText());
              } else if (cells.get(k).getText().equalsIgnoreCase("площадь:")) {
                excelDocument.setArea(cells.get(k + 1).getText());
              } else if (cells.get(k).getText().equalsIgnoreCase("Кадастровая стоимость, руб.:")) {
                excelDocument.setSum(cells.get(k + 1).getText());
              } else if (cells.get(k).getText().startsWith("Кадастровые номера расположенных в пределах")) {
                excelDocument.setNumbers(cells.get(k + 1).getText());
              } else if (cells.get(k).getText().equalsIgnoreCase("Правообладатель (правообладатели):")) {
                excelDocument.setOwner(cells.get(k + 2).getText());
              } else if (cells.get(k).getText().equalsIgnoreCase("Вид, номер и дата государственной регистрации права:")) {
                excelDocument.setInfo(cells.get(k + 2).getText());
              }
            }
          }
        });
      }
      pd.close();
      return excelDocument;
    } catch (IOException e) {
      log.error("Ошибка разбора pdf файла {}: {}", path, e);
      return null;
    }
  }

}
