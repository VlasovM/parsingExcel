import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class Runner {

    private static final String FILE_PATH = "./src/main/resources/"; // Пусть к файлу в вашей системе

    private static final String FILE_NAME = "ExcelTableTest.xlsx";

    private static final String FINAL_FILE_NAME = "FinalExcelTableTest.xlsx";


    public static void main(String[] args) throws IOException {
        var start = System.currentTimeMillis();
        var fileInputStream = new FileInputStream(FILE_PATH + FILE_NAME);
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheetAt(2); // Получить лист в таблице. (счёт с 0)
        System.out.println("Вы используете: " + sheet.getSheetName());

        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);

            // Сначала найдем колонку с ответами (например 23,24,1,4,5) - чтобы собрать ответы в коллекцию.
            List<String> answers = new ArrayList<>();
            Cell lastCellWithAnswers = row.getCell(row.getLastCellNum() - 1); // Получаем последний столбец, в котором по ТЗ лежаю перечисления ответов через ","
            // Когда лежат два значения X,Y - тип STRING, когда одно - X (если юзер дал один ответ) - то тип NUMERIC
            if (lastCellWithAnswers.getCellType().equals(CellType.STRING)) {
                String[] arrayAnswers = lastCellWithAnswers.getStringCellValue().split(",");
                answers = Arrays.asList(arrayAnswers);
            } else if (lastCellWithAnswers.getCellType().equals(CellType.NUMERIC)) {
                double doubleValue = lastCellWithAnswers.getNumericCellValue();
                int intValue = (int) doubleValue;
                answers.add(String.valueOf(intValue));
            }

            // Когда разделям ответы, то получится _23 или _20, где _ - пробел, избавимся от этого.
            answers = answers.stream()
                    .map(value -> value.replaceAll("\\s", ""))
                    .collect(Collectors.toList());

            // Затем идём по строке от начала до конца
            for (Cell cell : row) {

                if (!cell.getCellType().equals(CellType.BLANK)) {
                    // Игнорируем первый столбец и идём до столбца с ответами (но там ещё есть пустой столбец между последним ответом и ответами, поэтому -1 от ответов)
                    if (cell.getColumnIndex() != 0 || cell.getColumnIndex() != lastCellWithAnswers.getColumnIndex() - 1) {
                        var columnIndexCell = String.valueOf(cell.getColumnIndex());
                        if (answers.contains(columnIndexCell)) {
                            cell.setCellValue(1);
                        }
                    }
                } else {
                    break;
                }
            }
        }

        var fileOutputStream = new FileOutputStream(FILE_PATH + FINAL_FILE_NAME);
        workbook.write(fileOutputStream);
        workbook.close();

        System.out.println("Обработка завершена за: " + (System.currentTimeMillis() - start) + "ms");
    }
}
