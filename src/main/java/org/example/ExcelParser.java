package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.UUID;

public class ExcelParser {
    private static final String[] START_TIMES = {"090000", "105000", "124000", "154500"};
    private static final String[] END_TIMES = {"103500", "122500", "141500", "172000"};

    // Получаем текущий понедельник как начало недели
    private static final LocalDate START_DATE = getMonday(LocalDate.now()); // Начало недели — понедельник

    public static void main(String[] args) {
        String outputIcsPath = "D:\\calendar.ics";
        try {
            FileInputStream fis = new FileInputStream("D:\\timetable.xlsx");
            Workbook wb = new XSSFWorkbook(fis);
            BufferedWriter writer = new BufferedWriter(new FileWriter(outputIcsPath));
            {
                writer.write("BEGIN:VCALENDAR\n" +
                        "PRODID:-//Google Inc//Google Calendar 70.9054//EN\n" +
                        "VERSION:2.0\n" +
                        "CALSCALE:GREGORIAN\n" +
                        "METHOD:REQUEST\n" +
                        "X-WR-CALNAME:riza.seliamiyev@viti.edu.ua\n" +
                        "X-WR-TIMEZONE:Europe/Kiev\n");

                Sheet sheet = wb.getSheetAt(0); // Берем первый лист
                int startRow = 2;              // Начальная строка (например, строка 3 в Excel, индексация начинается с 0)

                // Проходим по всем дням
                for (int day = 0; day < 6; day++) {
                    LocalDate currentDate = START_DATE.plusDays(day);
                    String dateStr = currentDate.format(DateTimeFormatter.ofPattern("yyyyMMdd"));
                    int startCol = 1 + day * 8; // Первый столбец для текущего дня
                    int endCol = startCol + 8;  // Последний столбец для текущего дня

                    System.out.println("День " + (day + 1) + ":");


                    for (int pair = 0; pair < 4; pair++) {
                        int pairStartCol = startCol + pair * 2;
                        int pairEndCol = pairStartCol + 1;

                        // Получаем данные из ячеек для пары
                        Row row = sheet.getRow(startRow); // Строка для дисциплины
                        Cell disciplineCell = row.getCell(pairStartCol); // Название дисциплины
                        Cell lessonNumberCell = row.getCell(pairEndCol); // Номер занятия

                        // Извлекаем данные из ячеек
                        String discipline = (disciplineCell != null) ? disciplineCell.getStringCellValue().trim() : "";
                        String lessonNumber = (lessonNumberCell != null) ? lessonNumberCell.getStringCellValue().trim() : "";

                        // Если обе ячейки пустые, выводим "СЗСУ"
                        if (discipline.isEmpty() && lessonNumber.isEmpty()) {
                            discipline = "СЗСУ";
                            lessonNumber = "";
                        }
                        String summary = discipline;
                        if (!lessonNumber.isEmpty()) {
                            summary += " - " + lessonNumber;
                        }
                        String dtStart = dateStr + "T" + START_TIMES[pair];
                        String dtEnd = dateStr + "T" + END_TIMES[pair];
                        writer.write("BEGIN:VEVENT\n");
                        writer.write("DTSTART;TZID=Europe/Kiev:" + dtStart + "\n");
                        writer.write("DTEND;TZID=Europe/Kiev:" + dtEnd + "\n");
                        writer.write("ORGANIZER:mailto:riza.seliamiyev@viti.edu.ua\n");
                        writer.write("UID:" + generateUID() + "@temp.calendar.com"+ "\n");
                        writer.write("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;X-NUM-GUESTS=0:mailto:riza.seliamiyev@viti.edu.ua\n");
                        writer.write("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=INVITED;X-NUM-GUESTS=0:mailto:304ng@viti.edu.ua\n");
                        writer.write("SEQUENCE:0\n");
                        writer.write("STATUS:CONFIRMED\n");
                        writer.write("SUMMARY:" + summary + "\n");
                        writer.write("TRANSP:OPAQUE\n");
                        writer.write("END:VEVENT\n");
                    }
                }
                writer.write("END:VCALENDAR\n");
                writer.close();
                fis.close();
                System.out.println("ICS файл успешно создан: " + outputIcsPath);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    // Генерация уникального идентификатора для события
    private static String generateUID() {
        String uuid = UUID.randomUUID().toString().replace("-", "");
        // Обрезаем до первых 26 символов залупу себе обреж
        return uuid;
    }

    private static LocalDate getMonday(LocalDate date) {
        // Найти понедельник текущей недели
        return date.with(DayOfWeek.MONDAY);
    }
}
