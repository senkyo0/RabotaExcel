import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.data.category.DefaultCategoryDataset;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Главный класс приложения для обработки данных из Excel и создания статистики и графика.
 */
public class main1 {

    /**
     * Основной метод приложения, запускающий обработку данных и создающий статистику и график.
     *
     * @param args Аргументы командной строки (не используются).
     */
    public static void main(String[] args) {
        String inputFilePath = "input.xlsx";
        String outputFilePath = "output.xlsx";

        try {
            FileInputStream fis = new FileInputStream(new File(inputFilePath));
            Workbook workbook = WorkbookFactory.create(fis);

            Sheet sheet = workbook.getSheetAt(0);
            // Инициализация переменных для подсчета оценок и статистики
            int excellentCount = 0;
            int goodCount = 0;
            int passCount = 0;
            int failCount = 0;
            double totalScore = 0.0;
            int totalStudents = 0;
            int maxScore = 0;
            Map<Integer, List<String>> studentsByGrade = new HashMap<>();

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Пропускаем заголовок

                Cell nameCell = row.getCell(0);// Получаем ячейку с именем
                Cell scoreCell = row.getCell(1); // Получаем ячейку с оценкой

                if (nameCell != null && scoreCell != null) {
                    String name = nameCell.getStringCellValue(); // Имя студента
                    int score = (int) scoreCell.getNumericCellValue(); // Оценка
                    // Подсчет количества каждой оценки и общей статистики
                    switch (score) {
                        case 5:
                            excellentCount++;
                            break;
                        case 4:
                            goodCount++;
                            break;
                        case 3:
                            passCount++;
                            break;
                        case 2:
                            failCount++;
                            break;
                        default:
                            break;
                    }

                    totalScore += score; // Суммирование оценок
                    totalStudents++; // Подсчет общего числа студентов


                    if (score > maxScore) {
                        maxScore = score; // Определение максимальной оценки
                    }
                    // Группировка студентов по оценкам
                    if (!studentsByGrade.containsKey(score)) {
                        studentsByGrade.put(score, new ArrayList<>());
                    }

                    List<String> studentsList = studentsByGrade.get(score);
                    studentsList.add(name);
                }
            }
            // Расчет средней оценки
            double averageScore = totalStudents > 0 ? totalScore / totalStudents : 0;
            // Получаем или создаем лист "Результаты"
            Sheet resultSheet;
            if (workbook.getNumberOfSheets() > 0) {
                resultSheet = workbook.getSheet("Результаты");
            } else {
                resultSheet = workbook.createSheet("Результаты");
            }
            // Записываем результаты в уже существующий лист "Результаты"
            writeResultToSheet(resultSheet, 0, "Отличники", excellentCount);
            writeResultToSheet(resultSheet, 1, "Хорошисты", goodCount);
            writeResultToSheet(resultSheet, 2, "Троешники", passCount);
            writeResultToSheet(resultSheet, 3, "Не допущены", failCount);
            Row averageRow = resultSheet.createRow(4);
            averageRow.createCell(0).setCellValue("Средний балл");
            Cell averageCell = averageRow.createCell(1);
            averageCell.setCellValue(averageScore);
            averageCell.setCellType(CellType.NUMERIC);
            // Записываем ФИО студентов с определенными оценками
            int rowIndex = 5;
            for (Map.Entry<Integer, List<String>> entry : studentsByGrade.entrySet()) {
                List<String> studentsList = entry.getValue();
                if (studentsList != null) {
                    writeStudentsToSheet(resultSheet, rowIndex++, "Оценка " + entry.getKey(), studentsList);
                }
            }
            // Создание данных для графика
            DefaultCategoryDataset dataset = new DefaultCategoryDataset();
            dataset.addValue(excellentCount, "Количество", "Отличники");
            dataset.addValue(goodCount, "Количество", "Хорошисты");
            dataset.addValue(passCount, "Количество", "Троешники");
            dataset.addValue(failCount, "Количество", "Не допущены");
            // Создание графика
            JFreeChart chart = ChartFactory.createLineChart(
                    "Распределение оценок", // Заголовок графика
                    "Группы", // Метка оси X
                    "Количество", // Метка оси Y
                    dataset, // Данные для графика
                    org.jfree.chart.plot.PlotOrientation.VERTICAL,
                    true, true, false
            );
            // Сохранение графика в файле PNG
            File chartFile = new File("chart.png");
            ChartUtils.saveChartAsPNG(chartFile, chart, 800, 600);
            // Добавляем максимальную оценку в лист "Результаты"
            Row maxScoreRow = resultSheet.createRow(resultSheet.getLastRowNum() + 2);
            maxScoreRow.createCell(0).setCellValue("Максимальная оценка");
            maxScoreRow.createCell(1).setCellValue(maxScore);
            // Запись результатов и графика в файл Excel
            FileOutputStream fos = new FileOutputStream(outputFilePath);
            workbook.write(fos);
            fos.close();

            System.out.println("Результаты успешно записаны в файл " + outputFilePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Метод для записи результатов в лист Excel.
     *
     * @param sheet    Лист, в который будут записаны результаты.
     * @param rowIndex Индекс строки, в которую будут записаны результаты.
     * @param header   Заголовок для записи.
     * @param value    Значение для записи.
     */
    public static void writeResultToSheet(Sheet sheet, int rowIndex, String header, int value) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        row.createCell(0).setCellValue(header);
        row.createCell(1).setCellValue(value);
    }

    /**
     * Метод для записи данных о студентах в лист Excel.
     *
     * @param sheet        Лист, в который будут записаны данные о студентах.
     * @param rowIndex     Индекс строки, в которую будут записаны данные о студентах.
     * @param header       Заголовок для записи.
     * @param studentsList Список студентов для записи.
     */
    public static void writeStudentsToSheet(Sheet sheet, int rowIndex, String header, List<String> studentsList) {
        Row row = sheet.getRow(rowIndex); // Получение строки по индексу rowIndex из переданного листа
        if (row == null) { // Проверка, существует ли строка по указанному индексу, если нет, то создаем новую строку
            row = sheet.createRow(rowIndex);
        }
        row.createCell(0).setCellValue(header); // Создание ячейки в строке с индексом 0 и запись в нее значения заголовка
        int cellIndex = 1; // Инициализация переменной для отслеживания индекса ячейки
        for (String student : studentsList) {   // Запись имен студентов в ячейки строки
            row.createCell(cellIndex++).setCellValue(student);
        }
    }
}

