package org.example;

//public class ExcelParser {
//
//    // Временные интервалы для пар
//    private static final String[] START_TIMES = {"090000", "105000", "124000", "154500"};
//    private static final String[] END_TIMES = {"103500", "122500", "141500", "172000"};
//
//    // Даты для всех дней недели
//    private static final String[] WEEK_DAYS = {"20241216", "20241217", "20241218", "20241219", "20241220","20241221"};
//
//    public static void main(String[] args) {
//        String excelPath = "D:\\timetable.xlsx";
//        String outputIcsPath = "D:\\schedule.ics";
//
//        try (FileInputStream fis = new FileInputStream(excelPath);
//             Workbook workbook = new XSSFWorkbook(fis);
//             BufferedWriter writer = new BufferedWriter(new FileWriter(outputIcsPath))) {
//
//            // Инициализация файла календаря
//            writer.write("BEGIN:VCALENDAR\n" +
//                    "PRODID:-//Google Inc//Google Calendar 70.9054//EN\n" +
//                    "VERSION:2.0\n" +
//                    "CALSCALE:GREGORIAN\n" +
//                    "METHOD:REQUEST\n" +
//                    "X-WR-CALNAME:riza.seliamiyev@viti.edu.ua\n" +
//                    "X-WR-TIMEZONE:Europe/Kiev\n");
//
//            // Получаем первый лист
//            Sheet sheet = workbook.getSheetAt(0);
//
//            // Обрабатываем данные по всем дням недели
//            for (int day = 0; day < WEEK_DAYS.length; day++) {
//                int startCol = 1 + day * 8; // Столбец начала дня недели
//                int endCol = startCol + 8;  // Столбец конца дня недели
//
//                for (int rowIndex = 3; rowIndex <= 5; rowIndex++) { // Строки с парами
//                    Row row = sheet.getRow(rowIndex);
//                    writeEvent(writer, WEEK_DAYS[day], row, startCol, endCol);
//                }
//            }
//
//            // Завершаем файл календаря
//            writer.write("END:VCALENDAR\n");
//            System.out.println("ICS файл успешно создан: " + outputIcsPath);
//
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
//
//    private static void writeEvent(BufferedWriter writer, String date, Row row, int startCol, int endCol) throws IOException {
//        for (int colIndex = startCol; colIndex <= endCol; colIndex++) {
//            Cell cell = row.getCell(colIndex);
//            if (cell != null && cell.getCellType() == CellType.STRING) {
//                String summary = cell.getStringCellValue().trim();
//
//                if (!summary.isEmpty()) {
//                    int pairNumber = colIndex - startCol; // Номер пары
//                    String dtStart = date + "T" + START_TIMES[pairNumber]; // Начало
//                    String dtEnd = date + "T" + END_TIMES[pairNumber];     // Конец
//
//                    writer.write("BEGIN:VEVENT\n");
//                    writer.write("DTSTART;TZID=Europe/Kiev:" + dtStart + "\n");
//                    writer.write("DTEND;TZID=Europe/Kiev:" + dtEnd + "\n");
//                    writer.write("ORGANIZER:mailto:riza.seliamiyev@viti.edu.ua\n");
//                    writer.write("UID:" + generateUID() + "\n");
//                    writer.write("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;X-NUM-GUESTS=0:mailto:riza.seliamiyev@viti.edu.ua\n");
//                    writer.write("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;X-NUM-GUESTS=0:mailto:304ng@viti.edu.ua\n");
//                    writer.write("SEQUENCE:0\n");
//                    writer.write("STATUS:CONFIRMED\n");
//                    writer.write("SUMMARY:" + summary + "\n");
//                    writer.write("TRANSP:OPAQUE\n");
//                    writer.write("END:VEVENT\n");
//                }
//            }
//        }
//    }

//    private static String generateUID() {
//        return UUID.randomUUID().toString();
//    }
