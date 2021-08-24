package loadData;

import model.DataTK;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class ReportGenerator {

    private ArrayList<DataTK> allDataTK;
    private String fileName;
    private HashMap<String, ArrayList<DataTK>> separatedData;
    private ArrayList<DataTK> cleanedListTK;
    private ArrayList<DataTK> planTKList;
    private Map<String, String> listOfTesters;
    private Map<String, Integer> keyColumns = new HashMap<>();
    //private final String DATE_CONST = "yyyy-mm-nn";

    public ReportGenerator(String fileName, Map<String, Integer> keyColumns) throws IOException, IndexOutOfBoundsException, ParseException, FileNotFoundException {
        this.fileName = fileName;
        this.keyColumns = keyColumns;
        if (fileName.endsWith("xlsx")) {
            loadDataXLSX();
        }
        if (fileName.endsWith("csv")) {
            loadData();
        }
        separation();
        makeCleanSortedList();
    }

    //На будущее, в адаптависте нет CSV, можно вообще выпилить 
    private void loadData() throws IOException, IndexOutOfBoundsException, ParseException {
        this.allDataTK = new ArrayList<>();
        this.listOfTesters = new HashMap<>();
        try (BufferedReader bufferedReader = new BufferedReader(new FileReader(this.fileName))) {
            bufferedReader.readLine();
            while (bufferedReader.ready()) {
                String line = bufferedReader.readLine();
                String[] split = line.split(";");
                //костыль, надо убрать
                if (split.length < 20) {
                    continue;
                }
                this.allDataTK.add(addDataTKtoAll(split));
            }
        }
    }

    private void loadDataXLSX() throws IOException, IndexOutOfBoundsException, ParseException {
        FileInputStream fileInputStream = new FileInputStream(this.fileName);
        this.allDataTK = new ArrayList<>();
        this.listOfTesters = new HashMap<>();
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        this.keyColumns = findColumns(this.keyColumns, this.fileName);
        int k = sheet.getLastRowNum();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
//            System.out.println(row.getCell(0).getStringCellValue() + " " + i);
            DataTK dataTK = null;
            try {
                String[] split = new String[row.getLastCellNum()];
                for (int j = 0; j <= row.getLastCellNum(); j++) {
                    if (row.getCell(j) != null && row.getCell(j).getCellType() != CellType.BLANK) {
                        split[j] = row.getCell(j).getStringCellValue();
                    }
                }

                dataTK = addDataTKtoAll(split);
                this.allDataTK.add(dataTK);
            } catch (NegativeArraySizeException e) {
                continue;
            }
            System.out.println(dataTK.getStatus());
        }
    }

    //Впендюрить проверку на дату и сделать экспешн
    private DataTK addDataTKtoAll(String[] split) throws IndexOutOfBoundsException, ParseException {
        DataTK dataTK;
        String execution = this.keyColumns.get("Execution.Assigned To") > -1 ? split[this.keyColumns.get("Execution.Assigned To")] : "";
        String plannedStartDate = this.keyColumns.get("Execution.Planned Start Date") > -1 ? split[this.keyColumns.get("Execution.Planned Start Date")] : "";
        String plannedEndDate = this.keyColumns.get("Execution.Planned End Date") > -1 ? split[this.keyColumns.get("Execution.Planned End Date")] : "";
        String actualStartDate = this.keyColumns.get("Execution.Actual Start Date") > -1 ? split[this.keyColumns.get("Execution.Actual Start Date")] : "";
        String actualEndDate = this.keyColumns.get("Execution.Actual End Date") > -1 ? split[this.keyColumns.get("Execution.Actual End Date")] : "";
        String environment = this.keyColumns.get("Execution.Environment") > -1 ? split[this.keyColumns.get("Execution.Environment")] : "";
        String testCycle = this.keyColumns.get("Test Cycle.Key") > -1 ? split[this.keyColumns.get("Test Cycle.Key")] : "";
        String keyTK = this.keyColumns.get("Test Case.Key") > -1 ? split[this.keyColumns.get("Test Case.Key")] : "";
        String status = this.keyColumns.get("Execution.Result") > -1 ? split[this.keyColumns.get("Execution.Result")] : "";
        String nameTK = this.keyColumns.get("Test Case.Name") > -1 ? split[this.keyColumns.get("Test Case.Name")] : "";
        String owner = this.keyColumns.get("Test Cycle.Department") > -1 ? split[this.keyColumns.get("Test Cycle.Department")] : "";
        dataTK = new DataTK(execution, plannedStartDate, plannedEndDate, actualStartDate, actualEndDate, environment, testCycle, keyTK, nameTK, status, owner);
        return dataTK;
    }

    private boolean checkDateFormat(String date) throws ParseException {
        try {
            SimpleDateFormat sdf = new SimpleDateFormat("YYYY-mm-dd");
            sdf.parse(date);
        } catch (ParseException e) {
            return false;
        }
        return true;
    }

    private void separation() {
        this.separatedData = new HashMap<>();
        for (int i = 0; i < this.allDataTK.size(); i++) {
            if (!this.separatedData.containsKey(this.allDataTK.get(i).getKeyTK())) {
                this.separatedData.put(this.allDataTK.get(i).getKeyTK(), new ArrayList<>());
            }
            this.separatedData.get(this.allDataTK.get(i).getKeyTK()).add(this.allDataTK.get(i));
        }
    }

    //После GUI закомментить
    public void saveDataCSV(String fileName) throws IOException {
        try (BufferedWriter outputStream = new BufferedWriter(new FileWriter(fileName, Charset.forName("windows-1251")))) {
            outputStream.write("execution;keyTK;status;plannedStartDate;plannedEndDate;actualStartDate;actualEndDate;environment;testCycle;nameTK\n");
            for (int i = 0; i < this.cleanedListTK.size(); i++) {
                String buf = "";
                DataTK dataTK = this.cleanedListTK.get(i);
                buf += dataTK.getExecution() + ";" + dataTK.getKeyTK() + ";" + dataTK.getStatus() + ";" + dataTK.getPlannedStartDate() + ";" + dataTK.getPlannedEndDate()
                        + ";" + dataTK.getActualStartDate()
                        + ";" + dataTK.getActualEndDate() + ";" + dataTK.getEnvironment() + ";" + dataTK.getTestCycle()
                        + ";" + dataTK.getNameTK() + "\n";
                outputStream.write(buf);
            }
        }
    }

    private void makeCleanSortedList() {
        this.cleanedListTK = new ArrayList<>();
        for (Map.Entry<String, ArrayList<DataTK>> entry : this.separatedData.entrySet()) {
            DataTK dataTK = searchActualTK(entry.getValue());
            this.cleanedListTK.add(dataTK);
        }
        this.cleanedListTK.sort((o1, o2) -> {
            try {
                Date startDateTK1 = new SimpleDateFormat("dd.MM.yyyy").parse(o1.getPlannedStartDate());
                Date startDateTK2 = new SimpleDateFormat("dd.MM.yyyy").parse(o2.getPlannedStartDate());
                return startDateTK1.compareTo(startDateTK2);
            } catch (ParseException | NullPointerException e) {
                e.getMessage();
            }
            return 0;
        });
    }

    //Присваиваем статусную модель в соответствии с потребонстиями в отчете. Статус "Блокирован" равноценен проваленному кейсу
    private DataTK searchActualTK(ArrayList<DataTK> listTK) {
        for (int i = 0; i < listTK.size(); i++) {
            if (listTK.get(i).getStatus().equals("Pass")) {
                return listTK.get(i);
            }
        }
        for (int i = 0; i < listTK.size(); i++) {
            if (listTK.get(i).getStatus().equals("Blocked")) {
                listTK.get(i).setStatus("Fail");
                return listTK.get(i);
            }
        }
        for (int i = 0; i < listTK.size(); i++) {
            if (listTK.get(i).getStatus().equals("Fail")) {
                return listTK.get(i);
            }
        }
        for (int i = 0; i < listTK.size(); i++) {
            if (listTK.get(i).getStatus().equals("Not Executed")) {
                return listTK.get(i);
            }
        }
        return new DataTK();
    }

    //Не самый актуальный метод для плана на день
    private void planDay(String startDate, String endDate) throws ParseException {
        this.planTKList = new ArrayList<>();
        Date start = new SimpleDateFormat("yyyy-MM-dd").parse(startDate);
        Date end = new SimpleDateFormat("yyyy-MM-dd").parse(endDate);
        for (int i = 0; i < this.cleanedListTK.size(); i++) {
            Date startDateTK = new SimpleDateFormat("yyyy-MM-dd").parse(this.cleanedListTK.get(i).getPlannedStartDate());
            if (startDateTK.after(start) && startDateTK.before(end) || startDateTK.equals(start) || startDateTK.equals(end)) {
                this.planTKList.add(this.cleanedListTK.get(i));
            }
        }
        this.planTKList.sort((o1, o2) -> {
            try {
                Date startDateTK1 = new SimpleDateFormat("yyyy-MM-dd").parse(o1.getPlannedStartDate());
                Date startDateTK2 = new SimpleDateFormat("yyyy-MM-dd").parse(o2.getPlannedStartDate());
                return startDateTK1.compareTo(startDateTK2);
            } catch (ParseException e) {
                System.out.println(e.getMessage());
            }
            return 0;
        });
    }

    //Метод предназначен для формирования плана CSV формата. Содержит набор кейсов. Можно сильно улучшить в будущем, например через мапы
    //Подумать насчет массива с фамилиями тестировщиков
    //Изменить входные типы дат?
    public void makePlanTodayCSV(String startDate, String endDate, String fileName) throws IOException, ParseException {
        planDay(startDate, endDate);
        try (BufferedOutputStream outputStream = new BufferedOutputStream(new FileOutputStream(fileName))) {
            StringBuilder buf = new StringBuilder("Ключ;Название ТК;Planned Start Date;Planned End Date;Actual Start Date;Actual"
                    + " End Date;Среда;Конечный статус;" + generateNamesColumns() + "\n");
            for (DataTK dataTK : this.planTKList) {
                buf.append(dataTK.getKeyTK()).append(";").append(dataTK.getNameTK()).append(";").
                        append(dataTK.getPlannedStartDate()).append(";").append(dataTK.getPlannedEndDate()).append(";").
                        append(dataTK.getActualStartDate()).append(";").append(dataTK.getActualEndDate()).append(";").
                        append(dataTK.getEnvironment()).append(";").append(statusConvert(dataTK.getStatus())).append(";")
                        .append(generateStatusColumns(dataTK)).append("\n");
            }
            byte[] bytes = buf.toString().getBytes("windows-1251");
            outputStream.write(bytes);
        }
    }

    //Не реаилизовано
    //В ГУИ НАДО ЗАПОЛНЯТЬ!
    private String generateNamesColumns() {
        StringBuilder names = new StringBuilder();
        for (Map.Entry<String, String> entry : this.listOfTesters.entrySet()) {
            String name = entry.getValue().equals("") ? entry.getKey() : entry.getValue();
            names.append("Статус ").append(name).append(";");
        }
        return names.toString();
    }

    //Не реализовано
    private String generateStatusColumns(DataTK dataTK) {
        StringBuilder statusColumns = new StringBuilder();
        for (Map.Entry<String, String> entry : this.listOfTesters.entrySet()) {
            statusColumns.append(statusConvert(searchStatus(dataTK.getKeyTK(), entry.getKey()))).append(";");
        }
        return statusColumns.toString();
    }

    //????
    private String searchStatus(String keyTK, String fio) {
        for (int i = 0; i < this.separatedData.get(keyTK).size(); i++) {
            if (fio.equals(this.separatedData.get(keyTK).get(i).getExecution())) {
                return this.separatedData.get(keyTK).get(i).getStatus();
            }
        }
        return "";
    }

    //Возможно расширение статусов, требует обновления или задание статусной модели с приоритетами обработки пользователем
    private String statusConvert(String status) {
        if (status.equals("Pass")) {
            return "Пройден";
        }
        if (status.equals("Not Executed")) {
            return "Не завершен";
        }
        if (status.equals("Fail")) {
            return "Провален";
        }
        if (status.equals("In Progress")) {
            return "В работе";
        }
        return "";
    }

    public void report(String startDateTK, String endDateTK, String today, String name) throws ParseException, IOException {
        ReportXLSX reportXLSX = new ReportXLSX(new ArrayList<>(this.cleanedListTK), this);
        reportXLSX.generateGraphicSheet(startDateTK, endDateTK, today, name);
    }

    //Сделать это через Cleaned List, чтобы статусы были актуальные и исключились дубли кейсов
    public HashMap<String, ArrayList<DataTK>> getAllTKByOwners() {
        HashMap<String, ArrayList<DataTK>> mapByOwners = new HashMap<>();
        for (int i = 0; i < this.cleanedListTK.size(); i++) {
            if (!mapByOwners.containsKey(this.cleanedListTK.get(i).getOwner())) {
                mapByOwners.put(this.cleanedListTK.get(i).getOwner(), new ArrayList<>());
            }          
            mapByOwners.get(this.cleanedListTK.get(i).getOwner()).add(this.cleanedListTK.get(i));
        }
        return mapByOwners;
    }

    //Не реализовано
    public Map<String, String> getListOfTesters() {
        return new HashMap(listOfTesters);
    }

    //Блок статических методов, предназначенный для подготовки файла для работы с этим классом , можно вынести в отдельный класс
    //Возможно стоит его убрать и оставить метод получения колонок файла
    public static Map<String, Integer> findColumns(Map<String, Integer> keyColumns, String fileName) throws FileNotFoundException, IOException {
        try {
            FileInputStream fileInputStream = new FileInputStream(fileName);
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);
            Row rowZero = sheet.getRow(0);
            for (int i = 0; i < rowZero.getLastCellNum(); i++) {
                Cell cell = rowZero.getCell(i);
                if (keyColumns.containsKey(cell.getStringCellValue())) {
                    keyColumns.put(cell.getStringCellValue(), i);
                }
            }
        } catch (NullPointerException e) {
            throw e;
        }
        return keyColumns;
    }

    //Получаем список колонок файла из первой строки
    public static ArrayList<String> getFileColumns(String fileName) throws FileNotFoundException, IOException {
        ArrayList<String> listColumns = new ArrayList<>();
        FileInputStream fileInputStream = new FileInputStream(fileName);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row rowZero = sheet.getRow(0);
        for (int i = 0; i < rowZero.getLastCellNum(); i++) {
            Cell cell = rowZero.getCell(i);
            listColumns.add(cell.getStringCellValue());
        }
        return listColumns;
    }

    //Проверяем внутреннее содержимое тест кейсов на наличие корректного содержания (формат дат, наличие заполненных полей)
    public static String checkFileDataXLSX(String dateFormat, ArrayList<Integer> notNulls, ArrayList<Integer> dateFormats, String fileName) throws IOException {
        String errors = "";
        FileInputStream fileInputStream = new FileInputStream(fileName);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            for (int j = 0; j < notNulls.size(); j++) {
                if ((row.getCell(notNulls.get(j)) == null || row.getCell(notNulls.get(j)).getCellType() == CellType.BLANK)) {
                    errors += "В строке " + (i + 1) + " не заполнено обязательное поле " + sheet.getRow(0).getCell(notNulls.get(j)).getStringCellValue() + "\n";
                }
            }
            for (int j = 0; j < dateFormats.size(); j++) {
                SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
                if ((row.getCell(dateFormats.get(j)) != null) && row.getCell(dateFormats.get(j)).getCellType() != CellType.BLANK) {
                    try {
                        int it = dateFormats.get(j);
                        String dateCell = row.getCell(it).toString();
                        sdf.parse(dateCell);
                    } catch (ParseException ex) {
                        errors += "В строке " + (i + 1) + " формат даты не соответствует формату YYYY-mm-dd " + sheet.getRow(0).getCell(dateFormats.get(j)).getStringCellValue() + "\n";
                    }
                }
            }
            fileInputStream.close();
        }

        return errors;
    }

    //Обновляем хедер в файле в соответствии с нужным форматом, возможно стоит перепилить, чтобы программа подстраивалась под файл
    public static void updateFile(String fileName, Map<String, String> mappedSourceToTargetColumns) throws FileNotFoundException, IOException {
        Workbook workbook = new XSSFWorkbook(new FileInputStream(fileName));
        try (FileOutputStream outputStream = new FileOutputStream(new File(fileName))) {
            Sheet sheet = workbook.getSheetAt(0);
            Row header = sheet.getRow(0);
            for (int i = 0; i < header.getLastCellNum(); i++) {
                Cell cell = header.getCell(i);
                if (mappedSourceToTargetColumns.containsKey(cell.getStringCellValue())) {
                    cell.setCellValue(mappedSourceToTargetColumns.get(cell.getStringCellValue()));
                }
            }
            workbook.write(outputStream);
        }
    }

    //Сделать константу формата дат в классе
    public static void correctDatesFormatsXLSX(String fileName, ArrayList<Integer> columnNumbers) throws IOException {
        Workbook workbook = new XSSFWorkbook(new FileInputStream(fileName));
        try (FileOutputStream outputStream = new FileOutputStream(new File(fileName))) {
            Sheet sheet = workbook.getSheetAt(0);
            SimpleDateFormat sdfTarget = new SimpleDateFormat("yyyy-MM-dd");
            for (int i = 0; i < columnNumbers.size(); i++) {
                for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                    Cell cell = sheet.getRow(j).getCell(columnNumbers.get(i));
                    //Важнее сделать хоть какой-то отчет, чем выхватить ошибку
                    try {
                        Date date = cell.getDateCellValue();
                        cell.setCellValue(sdfTarget.format(date));
                    } catch (IllegalStateException | NullPointerException e) {
                        continue;
                    }
                }
            }
            workbook.write(outputStream);
        }
    }
}
