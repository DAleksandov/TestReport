package program;

import GUI.MakeReportForm;
import loadData.ReportGenerator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.UnsupportedLookAndFeelException;

public class Program {
    public static void main(String[] args) {
           try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Windows".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MakeReportForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MakeReportForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MakeReportForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MakeReportForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new MakeReportForm().setVisible(true);
            }
        });
         
      /*  FileInputStream fileInputStream = new FileInputStream("Тестовые прогоны бизнеса регресс.xlsx");
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        System.out.println(sheet.getLastRowNum());
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            System.out.println(sheet.getRow(i).getLastCellNum());
        }


        ReportGenerator reportGenerator = new ReportGenerator("Тестовые прогоны бизнеса регресс.xlsx");
        reportGenerator.saveDataCSV("Data.csv");
        //yyyy-MM-dd
        reportGenerator.makePlanTodayCSV("2020-05-01", "2020-06-10", "TESTsaveCSVPlanToday.csv");
        reportGenerator.report("2020-06-01", "2020-06-31", "2020-06-16");*/

        //Модуль маппинга сделать для загрузки из любого плагина+
        //Форматирование дат
        //Проверить пропертис отдельно и сделать их для всего проекта
        //Отчет по департаментам
        
        //Сделать график дня сегодняшнего
        //Сделать формирование прохождения за день
        //Считать количество не загруженных кейсов и выдавать список этих записей
    }
}
