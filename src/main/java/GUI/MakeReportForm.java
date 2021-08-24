package GUI;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.prefs.Preferences;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import loadData.ReportGenerator;
import org.apache.poi.EmptyFileException;

public class MakeReportForm extends javax.swing.JFrame {

    private File choosedFile;
    private static final String LAST_USED_FOLDER_CHOSE = "last_user_folder";
    private static final String LAST_USED_FOLDER_SAVE = "last_user_folder";
    private Map<String, Integer> keyColumns = new HashMap<>();
    private ArrayList<String> fileColumns;

    //Надо вынести в заполненный файл, потом апдейтить его
    {
        this.keyColumns.put("Execution.Assigned To", -1);
        this.keyColumns.put("Execution.Result", -1);
        this.keyColumns.put("Execution.Planned Start Date", -1);
        this.keyColumns.put("Execution.Actual Start Date", -1);
        this.keyColumns.put("Execution.Actual End Date", -1);
        this.keyColumns.put("Execution.Planned End Date", -1);
        this.keyColumns.put("Execution.Environment", -1);
        this.keyColumns.put("Test Cycle.Key", -1);
        this.keyColumns.put("Test Case.Key", -1);
        this.keyColumns.put("Test Case.Name", -1);
        this.keyColumns.put("Test Cycle.Department", -1);
        this.keyColumns.put("Test Case.Owner", -1);
    }

    public MakeReportForm() {
        this.fileColumns = new ArrayList<String>();
        this.choosedFile = null;
        File file = new File("\\Tools\\Preferences");
        initComponents();
    }


    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jDatePicker1 = new org.jdatepicker.JDatePicker();
        button2 = new java.awt.Button();
        jPanel1 = new javax.swing.JPanel();
        button1 = new java.awt.Button();
        makeReport = new java.awt.Button();
        makePlanToday = new java.awt.Button();
        label2 = new java.awt.Label();
        label4 = new java.awt.Label();
        jDatePicker3 = new org.jdatepicker.JDatePicker();
        jDatePicker4 = new org.jdatepicker.JDatePicker();
        jDatePicker5 = new org.jdatepicker.JDatePicker();
        label5 = new java.awt.Label();
        checkbox1 = new java.awt.Checkbox();
        jMenuBar2 = new javax.swing.JMenuBar();

        button2.setLabel("Задать имена");
        button2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                button2ActionPerformed(evt);
            }
        });

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setResizable(false);

        button1.setLabel("Выбрать файл");
        button1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                button1ActionPerformed(evt);
            }
        });

        makeReport.setLabel("Сформировать отчет");
        makeReport.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                makeReportActionPerformed(evt);
            }
        });

        makePlanToday.setLabel("Сформировать план");
        makePlanToday.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                makePlanTodayActionPerformed(evt);
            }
        });

        label2.setText("Начало");

        label4.setText("Завершение");

        jDatePicker3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jDatePicker3ActionPerformed(evt);
            }
        });

        jDatePicker5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jDatePicker5ActionPerformed(evt);
            }
        });

        label5.setText("Красная граница");

        checkbox1.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        checkbox1.setLabel("Использовать прошлый маппинг");
        checkbox1.setState(true);

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(0, 0, Short.MAX_VALUE)
                        .addComponent(checkbox1, javax.swing.GroupLayout.PREFERRED_SIZE, 214, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(label4, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(label2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(label5, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jDatePicker4, javax.swing.GroupLayout.DEFAULT_SIZE, 136, Short.MAX_VALUE)
                            .addComponent(jDatePicker3, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE)
                            .addComponent(jDatePicker5, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE))
                        .addGap(23, 23, 23)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(makePlanToday, javax.swing.GroupLayout.PREFERRED_SIZE, 199, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(makeReport, javax.swing.GroupLayout.PREFERRED_SIZE, 199, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(button1, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addGap(32, 32, 32))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(32, 32, 32)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jDatePicker4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(makeReport, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(label2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jDatePicker3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(label4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(makePlanToday, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE))
                .addGap(16, 16, 16)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                    .addComponent(jDatePicker5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(label5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(button1, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 20, Short.MAX_VALUE)
                .addComponent(checkbox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
        );

        setJMenuBar(jMenuBar2);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGap(94, 94, 94))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 16, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void button2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_button2ActionPerformed
        if (this.choosedFile == null) {
            JOptionPane.showMessageDialog(rootPane, "Сначала необходимо выбрать файл", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
    }//GEN-LAST:event_button2ActionPerformed

    private void jDatePicker5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jDatePicker5ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jDatePicker5ActionPerformed

    private void jDatePicker3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jDatePicker3ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jDatePicker3ActionPerformed

    //Сформировать план на день, маловероятна актуальность
    private void makePlanTodayActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_makePlanTodayActionPerformed
        String[] dates = checkDates();
        if (dates == null) {
            return;
        }
        JFileChooser chooser = new JFileChooser();
        chooser.setSelectedFile(new File(".csv"));
        chooser.showSaveDialog(this);
        try {
            ReportGenerator reportGenerator = new ReportGenerator(this.choosedFile.getAbsolutePath(), this.keyColumns);
            reportGenerator.saveDataCSV("logData" + dates[2] + ".csv");
            reportGenerator.makePlanTodayCSV(dates[0], dates[1], chooser.getSelectedFile().getAbsolutePath());
        } catch (IOException | ParseException ex) {
            Logger.getLogger(MakeReportForm.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_makePlanTodayActionPerformed

    private void makeReportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_makeReportActionPerformed
        //Заполняем последний путь выборщика файлов
        Preferences prefs = Preferences.userRoot().node(getClass().getName());
        JFileChooser chooser = new JFileChooser(prefs.get(LAST_USED_FOLDER_SAVE,
                new File(".").getAbsolutePath()));
        chooser.setSelectedFile(new File(""));
        //Проверяем заполнения полей
        String[] dates = checkDates();
        if (dates == null) {
            return;
        }
        try {
            //Получаем несоответствие колонок файла и ключевых колонок, МОЖНО УПРОСТИТЬ
            this.keyColumns = ReportGenerator.findColumns(this.keyColumns, this.choosedFile.getAbsolutePath());
            //Получаем список колонок в файле
            this.fileColumns = ReportGenerator.getFileColumns(this.choosedFile.getAbsolutePath());
            //Проверяем наличие несовпадающих колонок.    
            if (checkColumns()) {
                //Заполняем массивы обязательных проверок для заполнения полей и правильного формата дат. Пользователю не хочется такое доверять, лучше хардкод.    
                ArrayList<Integer> notNulls = new ArrayList<>();
                //Переделать через енам
                notNulls.add(this.keyColumns.get("Execution.Planned Start Date"));
                notNulls.add(this.keyColumns.get("Execution.Planned End Date"));
                notNulls.add(this.keyColumns.get("Execution.Result"));
                notNulls.add(this.keyColumns.get("Test Case.Key"));
                ArrayList<Integer> dateFormats = new ArrayList<>();
                dateFormats.add(this.keyColumns.get("Execution.Planned Start Date"));
                dateFormats.add(this.keyColumns.get("Execution.Actual Start Date"));
                dateFormats.add(this.keyColumns.get("Execution.Actual End Date"));
                dateFormats.add(this.keyColumns.get("Execution.Planned End Date"));
                int result = chooser.showSaveDialog(this);
                //Обновляем последний путь для выборщика файлоы
                if (result == JFileChooser.APPROVE_OPTION) {
                    prefs.put(LAST_USED_FOLDER_SAVE, chooser.getSelectedFile().getParent());
                    //Проверяем содержимое файла и генерируем лист ошибок, если что-то не так
                    String err = ReportGenerator.checkFileDataXLSX("YYYY-mm-dd", notNulls, dateFormats, this.choosedFile.getAbsolutePath());
                    if (!err.equals("")) {
                        new FileErrForm(err).setVisible(true);
                    } else {
                        //Составляем отчет
                        ReportGenerator reportGenerator = new ReportGenerator(this.choosedFile.getAbsolutePath(), this.keyColumns);
                        reportGenerator.saveDataCSV("logData" + dates[2] + ".csv");
                        reportGenerator.report(dates[0], dates[1], dates[2], chooser.getSelectedFile().getAbsolutePath());
                        //Снести в админку в будущем,не реализовано
                        //MakeNameFile makeNameFile = new MakeNameFile();
                        //makeNameFile.setVisible(true);
                        JOptionPane.showMessageDialog(rootPane, "Отчет сформирован", "Success", JOptionPane.INFORMATION_MESSAGE);
                    }
                }
            }
        } catch (IOException | ParseException | EmptyFileException ex) {
            JOptionPane.showMessageDialog(rootPane, "Некорректный файл или его содержание", "Error", JOptionPane.ERROR_MESSAGE);
        }
    }//GEN-LAST:event_makeReportActionPerformed
    //Если есть несовпадающие колонки, то будет открыта форма корректировки файла и маппинга существующих
    private boolean checkColumns() {
        ArrayList<String> listBadColumns = new ArrayList<>();
        for (Map.Entry<String, Integer> entry : this.keyColumns.entrySet()) {
            if (entry.getValue() == -1) {
                listBadColumns.add(entry.getKey());
            }
        }
        if (!(listBadColumns.isEmpty() || checkbox1.getState())) {
            new MakeColumnsFileForm(this.choosedFile.getAbsolutePath(), this.keyColumns, this.fileColumns).setVisible(true);
            return false;
        }
        //Неудачная попытка прогружать прошлый шаблон
        /* if (!listBadColumns.isEmpty() || checkbox1.getState()) {
            FileInputStream fileInputStream = new FileInputStream(new File("Properties"));                  
            Properties properties = new Properties();
            properties.load(fileInputStream);
            Map<String, String> map = new HashMap<String, String>();
            map.
        }
         */
        return true;
    }

    private void button1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_button1ActionPerformed
        Preferences prefs = Preferences.userRoot().node(getClass().getName());
        JFileChooser chooser = new JFileChooser(prefs.get(LAST_USED_FOLDER_CHOSE,
                new File(".").getAbsolutePath()));
        chooser.setSelectedFile(new File(".xlsx"));
        FileNameExtensionFilter filterXLSX = new FileNameExtensionFilter("XLSX Files", "xlsx");
        FileNameExtensionFilter filterCSV = new FileNameExtensionFilter("CSV Files", "csv");
        chooser.removeChoosableFileFilter(chooser.getFileFilter());
        chooser.addChoosableFileFilter(filterXLSX);
        chooser.addChoosableFileFilter(filterCSV);
        chooser.setFileFilter(filterXLSX);
        int result = chooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            prefs.put(LAST_USED_FOLDER_CHOSE, chooser.getSelectedFile().getParent());
        }
        this.choosedFile = chooser.getSelectedFile();

    }//GEN-LAST:event_button1ActionPerformed

    //Наабор интерфейсных проверок и возврат строковых представлений дат (можно попробовать перепилить на Date)
    private String[] checkDates() {
        if (this.choosedFile == null) {
            JOptionPane.showMessageDialog(rootPane, "Сначала необходимо выбрать файл", "Error", JOptionPane.ERROR_MESSAGE);
            return null;
        }
        if (jDatePicker4.getModel().getValue() == null) {
            JOptionPane.showMessageDialog(rootPane, "Заполните дату начала", "Error", JOptionPane.ERROR_MESSAGE);
            return null;
        }
        if (jDatePicker3.getModel().getValue() == null) {
            JOptionPane.showMessageDialog(rootPane, "Заполните дату окончания", "Error", JOptionPane.ERROR_MESSAGE);
            return null;
        }
        //Толстый блок для проверки введенных значений
        SimpleDateFormat simple = new SimpleDateFormat("yyyy-MM-dd");
        GregorianCalendar grStart = new GregorianCalendar(jDatePicker4.getModel().getYear(), jDatePicker4.getModel().getMonth(), jDatePicker4.getModel().getDay());
        GregorianCalendar grEnd = new GregorianCalendar(jDatePicker3.getModel().getYear(), jDatePicker3.getModel().getMonth(), jDatePicker3.getModel().getDay());
        GregorianCalendar grtoday = new GregorianCalendar(jDatePicker5.getModel().getYear(), jDatePicker5.getModel().getMonth(), jDatePicker5.getModel().getDay());
        Date dateStart = new Date(grStart.getTimeInMillis());
        Date dateEnd = new Date(grEnd.getTimeInMillis());
        Date todayD = new Date(grtoday.getTimeInMillis());

        if (dateEnd.getTime() < dateStart.getTime()) {
            JOptionPane.showMessageDialog(rootPane, "Дата окончания не может быть меньше даты начала", "Error", JOptionPane.ERROR_MESSAGE);
            return null;
        }
        if (todayD.getTime() < dateStart.getTime() || todayD.getTime() > dateEnd.getTime()) {
            JOptionPane.showMessageDialog(rootPane, "Граница должна попадать во временной интервал между стартом и окончанием", "Error", JOptionPane.ERROR_MESSAGE);
            return null;
        }
        Date today = new Date(grtoday.getTimeInMillis());
        //Возвращаем в нужном формате строковые представления дат
        String[] dates = {simple.format(dateStart), simple.format(dateEnd), simple.format(today)};
        return dates;
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private java.awt.Button button1;
    private java.awt.Button button2;
    private java.awt.Checkbox checkbox1;
    private org.jdatepicker.JDatePicker jDatePicker1;
    private org.jdatepicker.JDatePicker jDatePicker3;
    private org.jdatepicker.JDatePicker jDatePicker4;
    private org.jdatepicker.JDatePicker jDatePicker5;
    private javax.swing.JMenuBar jMenuBar2;
    private javax.swing.JPanel jPanel1;
    private java.awt.Label label2;
    private java.awt.Label label4;
    private java.awt.Label label5;
    private java.awt.Button makePlanToday;
    private java.awt.Button makeReport;
    // End of variables declaration//GEN-END:variables
}
