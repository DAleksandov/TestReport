package loadData;

import model.DataTK;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Map;

public class ReportXLSX {

    private ArrayList<DataTK> cleanedList;
    private Workbook XLSXWorkbook;
    private ReportGenerator reportGenerator;

    public ReportXLSX(ArrayList<DataTK> cleanedList, ReportGenerator reportGenerator) {
        this.cleanedList = cleanedList;
        this.XLSXWorkbook = new XSSFWorkbook();
        this.reportGenerator = reportGenerator;
    }

    //Переделать класс - полем доложен быть репортгенератор, из него должны получать основные данные или вообще тут должны быть статические методы
    //Форматы дат изменить
    //Конечная дата нужна, сделать
    public void generateGraphicSheet(String dateStart, String dateEnd, String today, String name) throws ParseException, IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(name);
        Sheet graphicSheet = this.XLSXWorkbook.createSheet("Graphic");
        //Создаем шапку
        Row row = graphicSheet.createRow(0);
        Cell date = row.createCell(0);
        date.setCellValue("Date");
        Cell plannedExecutionCell = row.createCell(1);
        plannedExecutionCell.setCellValue("Planned Execution");
        Cell plannedPassRateCell = row.createCell(2);
        plannedPassRateCell.setCellValue("Planned Pass Rate");
        Cell plannedActPassRateCell = row.createCell(3);
        plannedActPassRateCell.setCellValue("Actual Execution Rate");
        Cell cell4 = row.createCell(4);
        cell4.setCellValue("ActPassRateCell");
        //Создаем первую строку
        Row rowCycle1 = graphicSheet.createRow(1);
        //тут записываем даты с начала до конца
        Cell dateCycle1 = rowCycle1.createCell(0);
        dateCycle1.setCellValue(dateStart);
        //тут записываем количество запланированного начала к этому дню
        Cell plannedExecutionRate1 = rowCycle1.createCell(1);
        plannedExecutionRate1.setCellValue(plannedExecutionRate(dateStart));
        //тут записываем количество запланированного начала к этому дню
        Cell plannedPassRate1 = rowCycle1.createCell(2);
        plannedPassRate1.setCellValue(plannedPassRate(dateStart));
        //тут записываем количество реального начала к этому дню
        Cell actualExecutionRate1 = rowCycle1.createCell(3);
        actualExecutionRate1.setCellValue(actualExecutionRate(dateStart));
        //тут записываем количество реального начала к этому дню
        Cell actualPassRate1 = rowCycle1.createCell(4);
        actualPassRate1.setCellValue(actualPassRate(dateStart));
        //Календарь для прибавления даты
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(sdf.parse(dateStart));
        String nextDay = sdf.format(calendar.getTime());
        int countRows = 0;
        for (int i = 2; sdf.parse(nextDay).getTime() < sdf.parse(dateEnd).getTime(); i++) {
            countRows++;
            calendar.add(Calendar.DATE, 1);
            nextDay = sdf.format(calendar.getTime());
            Row rowCycle = graphicSheet.createRow(i);
            //тут записываем даты с начала до конца
            Cell dateCycle = rowCycle.createCell(0);
            if (sdf.parse(nextDay).getTime() <= sdf.parse(dateEnd).getTime()) {
                dateCycle.setCellValue(nextDay);
            }
            //тут записываем количество запланированного начала к этому дню
            Cell plannedExecutionRate = rowCycle.createCell(1);
            plannedExecutionRate.setCellFormula(plannedExecutionRate(nextDay) + "+" + "B" + (i));
            //тут записываем количество запланированного начала к этому дню
            Cell plannedPassRate = rowCycle.createCell(2);
            plannedPassRate.setCellFormula(plannedPassRate(nextDay) + "+" + "C" + (i));
            //тут записываем количество реального начатых к этому дню
            Cell actualExecutionRate = rowCycle.createCell(3);
            actualExecutionRate.setCellFormula(actualExecutionRate(nextDay) + "+" + "D" + (i));
            //тут записываем количество реально законченных к этому дню
            Cell actualPassRate = rowCycle.createCell(4);
            actualPassRate.setCellFormula(actualPassRate(nextDay) + "+" + "E" + (i));
        }
        generateGraphic(fileOutputStream, today, countRows);
    }

    //Азбавиться от этой херни в аргументах
    private void generateGraphic(FileOutputStream stream, String today, int countRows) throws IOException {
        XSSFSheet sheet = (XSSFSheet) this.XLSXWorkbook.getSheet("Graphic");
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 8, 0, 25, 25);
        XSSFChart chart = drawing.createChart(anchor);
        //Вытаскиваем фон графика и делаем с ним кучу всего
        CTPlotArea plotArea = chart.getCTChart().getPlotArea();
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.BOTTOM);
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle("Период тестирования");
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle("Количество кейсов");
        //Делаем левую ось красивой
        int majorUnit = this.cleanedList.size() < 10 ? 1 : 5;
        leftAxis.setMajorUnit(majorUnit);//шаг деления
        leftAxis.setMinimum(0);
        leftAxis.setMaximum(this.cleanedList.size());
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        //убрать хардкод, вынести в отдельные методы
        int count = 0;
        sheet.getRow(2).getOutlineLevel();
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            count++;
            if (sheet.getRow(i).getCell(0).getStringCellValue().equals(today)) {
                break;
            }
        }
        int g = -1;
        sheet.getRow(0).createCell(5).setCellValue("TodayGraph");
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            if (sheet.getRow(i).getCell(0).getStringCellValue().equals(today)) {
                sheet.getRow(i).createCell(5).setCellValue(g);
                g = i + 1;
                sheet.getRow(i + 1).createCell(5).setCellValue(100000000);
                break;
            }
            sheet.getRow(i).createCell(5).setCellValue(g);
        }
        //Создаем сами линии на графике
        XDDFCategoryDataSource xs1 = XDDFDataSourcesFactory.fromStringCellRange((XSSFSheet) this.XLSXWorkbook.getSheet("Graphic"), new CellRangeAddress(1, countRows, 0, 0));
        XDDFCategoryDataSource xs2 = XDDFDataSourcesFactory.fromStringCellRange((XSSFSheet) this.XLSXWorkbook.getSheet("Graphic"), new CellRangeAddress(1, count, 0, 0));
        XDDFCategoryDataSource xs3 = XDDFDataSourcesFactory.fromStringCellRange((XSSFSheet) this.XLSXWorkbook.getSheet("Graphic"), new CellRangeAddress(1, g, 0, 0));
        XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, countRows, 1, 1));
        XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, countRows, 2, 2));
        XDDFNumericalDataSource<Double> ys3 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, count, 3, 3));
        XDDFNumericalDataSource<Double> ys4 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, count, 4, 4));
        XDDFNumericalDataSource<Double> todayGraph = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, g, 5, 5));

        XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
        XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(xs1, ys1);
        series1.setTitle("Planned Execution Rate", null); // https://stackoverflow.com/questions/21855842
        series1.setSmooth(false);

        // series1.setSmooth(false); // https://stackoverflow.com/questions/29014848
        series1.setMarkerStyle(MarkerStyle.NONE); // https://stackoverflow.com/questions/39636138
        XDDFLineChartData.Series series2 = (XDDFLineChartData.Series) data.addSeries(xs1, ys2);
        series2.setTitle("Planned Pass Rate", null);
        series2.setSmooth(false);
        //series2.setMarkerSize((short) 6);
        series2.setMarkerStyle(MarkerStyle.NONE); // https://stackoverflow.com/questions/39636138
        XDDFLineChartData.Series series3 = (XDDFLineChartData.Series) data.addSeries(xs2, ys3);
        series3.setTitle("Actual Execution Rate", null);
        series3.setSmooth(false);
        //  series3.setMarkerSize((short) 6);
        series3.setMarkerStyle(MarkerStyle.NONE); // https://stackoverflow.com/questions/39636138
        XDDFLineChartData.Series series4 = (XDDFLineChartData.Series) data.addSeries(xs2, ys4);
        series4.setTitle("Actual Pass Rate", null);
        series4.setSmooth(false);
        //  series4.setMarkerSize((short) 6);
        series4.setMarkerStyle(MarkerStyle.NONE); // https://stackoverflow.com/questions/39636138
        XDDFLineChartData.Series series5 = (XDDFLineChartData.Series) data.addSeries(xs3, todayGraph);
        series5.setTitle("Today", null);
        series5.setSmooth(false);
        chart.plot(data);
        // if your series have missing values like https://stackoverflow.com/questions/29014848
        // chart.displayBlanksAs(DisplayBlanks.GAP);
        // https://stackoverflow.com/questions/24676460
        solidLineSeries(data, 0, PresetColor.CORNFLOWER_BLUE, PresetLineDash.DASH);
        solidLineSeries(data, 1, PresetColor.LIGHT_GREEN, PresetLineDash.DASH);
        solidLineSeries(data, 2, PresetColor.CORNFLOWER_BLUE, PresetLineDash.SOLID);
        solidLineSeries(data, 3, PresetColor.LIGHT_GREEN, PresetLineDash.SOLID);
        solidLineSeries(data, 4, PresetColor.RED, PresetLineDash.SOLID);
        generateTableByDepartments(stream, today, today);
    }

    private static void solidLineSeries(XDDFChartData data, int index, PresetColor color, PresetLineDash dash) {
        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(color));
        XDDFLineProperties line = new XDDFLineProperties();
        line.setFillProperties(fill);
        line.setPresetDash(new XDDFPresetLineDash(dash));
        XDDFChartData.Series series = data.getSeries().get(index);
        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setLineProperties(line);
        series.setShapeProperties(properties);
    }

    //Считаем количество пройденных/заваленных кейсов
    private int plannedExecutionRate(String date) {
        int count = 0;
        for (int i = 0; i < this.cleanedList.size(); i++) {
            String startDateTK = this.cleanedList.get(i).getPlannedEndDate();
            if (date.equals(startDateTK)) {
                count++;
            }
        }
        return count;
    }

    private int plannedPassRate(String date) {
        int count = 0;
        for (int i = 0; i < this.cleanedList.size(); i++) {
            String startDateTK = this.cleanedList.get(i).getPlannedEndDate();
            if (date.equals(startDateTK)) {
                count++;
            }
        }
        return count;
    }

    private int actualExecutionRate(String date) {
        int count = 0;
        for (int i = 0; i < this.cleanedList.size(); i++) {
            DataTK dataTK = this.cleanedList.get(i);
            String startDateTK = this.cleanedList.get(i).getActualEndDate();
            if (date.equals(startDateTK) && (dataTK.getStatus().equals("Pass") || dataTK.getStatus().equals("Fail"))) {
                count++;
            }
        }
        return count;
    }

    private int actualPassRate(String date) {
        int count = 0;
        for (int i = 0; i < this.cleanedList.size(); i++) {
            DataTK dataTK = this.cleanedList.get(i);
            String startDateTK = this.cleanedList.get(i).getActualEndDate();
            if (date.equals(startDateTK) && dataTK.getStatus().equals("Pass")) {
                count++;
            }
        }
        return count;
    }

    private void generateTableByDepartments(FileOutputStream stream, String startDate, String endDate) throws IOException {
        //Делаем шапку
        //1 строка
        Sheet sheetT = this.XLSXWorkbook.createSheet();
        //Создаем стили для будущих ячеек
        CellStyle styleHeader = createBackGround(createBorders(this.XLSXWorkbook.createCellStyle()), this.XLSXWorkbook.createFont());
        CellStyle styleBorders = createBorders(this.XLSXWorkbook.createCellStyle());
        Row row = sheetT.createRow(0);
        Row row1 = sheetT.createRow(1);
        Cell c1 = row.createCell(0);
        c1.setCellStyle(styleHeader);
        c1.setCellValue("Департамент");
        Cell c2 = row1.createCell(0);
        c2.setCellStyle(styleHeader);
        sheetT.addMergedRegion(CellRangeAddress.valueOf("A1:A2"));
        Cell c3 = row.createCell(1);
        c3.setCellStyle(styleHeader);
        c3.setCellValue("План тест-кейсов " + startDate + " - " + endDate);
        row1.createCell(1).setCellStyle(styleHeader);
        sheetT.addMergedRegion(CellRangeAddress.valueOf("B1:B2"));
        Cell c4 = row.createCell(2);
        c4.setCellStyle(styleHeader);
        c4.setCellValue("Пройдено");
        Cell c5 = row.createCell(3);
        c5.setCellStyle(styleHeader);
        sheetT.addMergedRegion(CellRangeAddress.valueOf("C1:D1"));
        Cell c6 = row.createCell(4);
        c6.setCellStyle(styleHeader);
        c6.setCellValue("Осталось");
        //Cell c7 = row.createCell(5);
        //c7.setCellStyle(styleHeader);
        sheetT.addMergedRegion(CellRangeAddress.valueOf("E1:E2"));
        //2 строка
        Cell c8 = row1.createCell(2);
        c8.setCellStyle(styleHeader);
        c8.setCellValue("Без замечаний");
        Cell c9 = row1.createCell(3);
        c9.setCellValue("С замечаниями");
        c9.setCellStyle(styleHeader);
        int i = 2;
        //Создаем строки по владельцам прогонов со статистикой
        for (Map.Entry<String, ArrayList<DataTK>> entry : this.reportGenerator.getAllTKByOwners().entrySet()) {
            String key = entry.getKey();
            ArrayList<DataTK> value = entry.getValue();
            Row rowOwner = sheetT.createRow(i);
            rowOwner.createCell(0).setCellValue(key);
            rowOwner.getCell(0).setCellStyle(styleHeader);
            rowOwner.createCell(1).setCellValue(value.size());
            rowOwner.getCell(1).setCellStyle(styleBorders);
            int countPass = 0;
            int countFail = 0;
            for (int j = 0; j < value.size(); j++) {
                if (value.get(j).getStatus().equals("Pass")) {
                    countPass++;
                }
                if (value.get(j).getStatus().equals("Fail") || value.get(j).getStatus().equals("Blocked")) {
                    countFail++;
                }
            }
            rowOwner.createCell(2).setCellValue(countPass);
            rowOwner.getCell(2).setCellStyle(styleBorders);
            rowOwner.createCell(3).setCellValue(countFail);
            rowOwner.getCell(3).setCellStyle(styleBorders);
            rowOwner.createCell(4).setCellValue(value.size() - countFail - countPass);
            rowOwner.getCell(4).setCellStyle(styleBorders);
            i++;
        }
        for (int j = 0; j <= sheetT.getRow(0).getLastCellNum(); j++) {
            sheetT.autoSizeColumn(j, true);
        }
        this.XLSXWorkbook.write(stream);
        stream.close();
    }

    //Применяет выравнивание текста по центру яцейки + дает ей границы
    private CellStyle createBorders(CellStyle style) {
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    //Задает задний фор у ячейки + меняет цвет текста
    private CellStyle createBackGround(CellStyle style, Font font) {
        style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        return style;
    }

}
