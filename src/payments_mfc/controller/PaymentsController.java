package payments_mfc.controller;

import com.google.gson.Gson;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.collections.transformation.SortedList;
import javafx.concurrent.Task;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import payments_mfc.appController;
import payments_mfc.model.PaymentsModel;
import payments_mfc.model.PaymentsModelInput;
import payments_mfc.model.SettingsModel;

import java.awt.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.*;
import java.util.List;
import java.util.function.Predicate;

public class PaymentsController {


    // Класс для потока получения отчёта
    public static class ReportPayments extends Task<ArrayList<PaymentsModel>> {
        //private final LocalDate datePeriod; // Дата начала в юникс
        private final File file;


        public ReportPayments(File file) {
            this.file = file;
        }
        @Override
        protected ArrayList<PaymentsModel> call() throws Exception {
            // получение одного общего отчёта
            ArrayList<PaymentsModel> reportListFinal;
            reportListFinal=GetReportAll(file);
            // ArrayList<PaymentsModel> reportListFinal= GetReportAll(file);
            return reportListFinal; // Возвращаем итоговый отчёт
        }
    }
    public static ArrayList<PaymentsModel> GetReportAll(File file) throws IOException {

        List<PaymentsModel> paymentsList = new ArrayList<PaymentsModel>();
        if (file != null) {
            //Open
            System.out.println("Процесс открытия файла");


            // Read XSL file
            FileInputStream inputStream = new FileInputStream(file);

            // Get the workbook instance for XLS file
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

            // Get first sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            List<PaymentsModelInput> paymentsListInputTarrif= new ArrayList<PaymentsModelInput>();
            List<PaymentsModelInput> paymentsListInputNameMFC= new ArrayList<PaymentsModelInput>();
            List<PaymentsModelInput> paymentsListInputFioApplicant= new ArrayList<PaymentsModelInput>();
            List<PaymentsModelInput> paymentsListInputNameMFCAndTarriff= new ArrayList<PaymentsModelInput>();
            for (int cell_num=1; cell_num<=4; cell_num++){
                for (Row row : sheet) { // For each Row.
                    Cell cell = row.getCell(cell_num); // Get the Cell at the Index / Column you want.
                    CellType cellType = cell.getCellTypeEnum();
                    int number_of_list=1;
                    switch (cellType) {
                        case _NONE:
                            //System.out.print("");
                            //System.out.print("\t");
                            break;
                        case BOOLEAN:
                            //System.out.print(cell.getBooleanCellValue());
                            //System.out.print("\t");
                            break;
                        case BLANK:
                            //System.out.print("");
                            //System.out.print("\t");
                            break;
                        case FORMULA:
                            // Formula
                            //System.out.print(cell.getCellFormula());
                            //System.out.print("\t");

                            //FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                            // Print out value evaluated by formula
                            //System.out.print(evaluator.evaluate(cell).getNumberValue());
                            break;
                        case NUMERIC:
                            //System.out.print(cell.getNumericCellValue());
                            if (cell_num==1){
                                paymentsListInputTarrif.add(new PaymentsModelInput(cell.getNumericCellValue(),"",""));
                            }

                            break;
                        case STRING:
                            //System.out.print(cell.getStringCellValue());
                            //listFromExport.add(cell.getStringCellValue());
                            if (cell_num==1){
                                paymentsListInputTarrif.add(new PaymentsModelInput(0,"",""));
                            }
                            if (cell_num==2){
                                paymentsListInputNameMFC.add(new PaymentsModelInput(0,cell.getStringCellValue(),""));
                            }

                            if (cell_num==4){
                                paymentsListInputFioApplicant.add(new PaymentsModelInput(0,"",cell.getStringCellValue()));
                            }

                            //System.out.print("\n");
                            break;
                        case ERROR:
                            //System.out.print("!");
                            //System.out.print("\t");
                            break;
                    }
                }
            }
            System.out.println("Size mfc "+paymentsListInputNameMFC.size()+" Size tarrif "+ paymentsListInputTarrif.size());
            /*for (int i=0; i<paymentsListInputNameMFC.size(); i++){
                int numb=i+1;
                System.out.println("numb: "+numb+" "+ paymentsListInputTarrif.get(i).getName_mfc() +" "+ paymentsListInputTarrif.get(i).getTarriff());
                System.out.println("numb: "+numb+" "+ paymentsListInputNameMFC.get(i).getName_mfc() +" "+ paymentsListInputNameMFC.get(i).getTarriff());
            }*/

            for (int i=0; i<paymentsListInputTarrif.size(); i++){
                paymentsListInputNameMFCAndTarriff.add(new PaymentsModelInput(paymentsListInputTarrif.get(i).getTarriff(),paymentsListInputNameMFC.get(i).getName_mfc(),
                        paymentsListInputFioApplicant.get(i).getFio_applicant()));
            }

            /*for (int i=0; i< paymentsListInputNameMFCAndTarriff.size(); i++){
                int numb=i+1;
                System.out.println("numb: "+numb+" "+ paymentsListInputNameMFCAndTarriff.get(i).getTarriff() +" "+ paymentsListInputNameMFCAndTarriff.get(i).getName_mfc());
            }*/
            paymentsListInputNameMFCAndTarriff.remove(0);
            /*for (int i=0; i< paymentsListInputNameMFCAndTarriff.size(); i++){
                int numb=i+1;
                System.out.println("numb: "+numb+" "+ paymentsListInputNameMFCAndTarriff.get(i).getTarriff() +" "+ paymentsListInputNameMFCAndTarriff.get(i).getName_mfc());
            }*/
            List<String> mfcSwithDublicates=new ArrayList<>();
            List<String> mfcWithoutDublicates= new ArrayList<>();
            for(int i =0; i<paymentsListInputNameMFCAndTarriff.size(); i++){
                mfcSwithDublicates.add(paymentsListInputNameMFCAndTarriff.get(i).getName_mfc());
            }

            Set<String> setMfc = new HashSet<>(mfcSwithDublicates);
            mfcWithoutDublicates.addAll(setMfc);
            System.out.println("WITH DUBLICATES");
            for (int i=0; i< mfcSwithDublicates.size(); i++){
                int numb=i+1;
                System.out.println("numb "+numb+ " MFC "+mfcSwithDublicates.get(i));
            }
            System.out.println("WITHOUT DUBLICATES");
            for (int i=0; i< mfcWithoutDublicates.size(); i++){
                int numb=i+1;
                System.out.println("numb "+numb+ " MFC "+mfcWithoutDublicates.get(i));
            }

            int tariff_70_sum_allRegion=0;int tariff_110_sum_allRegion=0;int tariff_130_sum_allRegion=0;
            int tariff_220_sum_allRegion=0;int tariff_260_sum_allRegion=0;
            int tariff_310_sum_allRegion=0;int tariff_270_sum_allRegion=0;
            int tariff_total_region=0;
            for (int name_mfc=0; name_mfc<mfcWithoutDublicates.size();name_mfc++ ){
                int Sum_tariff_mfc=0;
                ObservableList<PaymentsModelInput> data = FXCollections.observableArrayList(paymentsListInputNameMFCAndTarriff);
                FilteredList<PaymentsModelInput> filteredData = new FilteredList<>(data, e -> true);
                int finalName_mfc = name_mfc;
                filteredData.setPredicate((Predicate<? super PaymentsModelInput>) paymentsModelInput->{
                    String lowerCaseFilter = mfcWithoutDublicates.get(finalName_mfc);
                    lowerCaseFilter=lowerCaseFilter.toLowerCase();
                    // Сравниваем обращения с фильтром
                    if (paymentsModelInput.getName_mfc().toLowerCase().contains(lowerCaseFilter)){
                        return true;
                    }
                    return false;
                });
                SortedList<PaymentsModelInput> sortedData = new SortedList<>(filteredData);

                System.out.println("SORTED DATA!!! \n");
                int tariff_70_sum=0;int tariff_110_sum=0;int tariff_130_sum=0;
                int tariff_220_sum=0;int tariff_260_sum=0;
                int tariff_310_sum=0;int tariff_270_sum=0;
                int tariff_other=0;
                String other_fio=""; String commentary="";

                for (int i=0; i<sortedData.size(); i++){
                    int numb=i+1;
                    Sum_tariff_mfc+= sortedData.get(i).getTarriff();
                    System.out.println("numb "+numb+ " Sorted "+ sortedData.get(i).getTarriff()+" "+sortedData.get(i).getName_mfc());

                    switch ((int) sortedData.get(i).getTarriff()){
                        case 70:
                            tariff_70_sum+=sortedData.get(i).getTarriff();
                            break;
                        case 110:
                            tariff_110_sum+=sortedData.get(i).getTarriff();
                            break;
                        case 130:
                            tariff_130_sum+=sortedData.get(i).getTarriff();
                            break;
                        case 220:
                            tariff_220_sum+=sortedData.get(i).getTarriff();
                            break;
                        case 260:
                            tariff_260_sum+=sortedData.get(i).getTarriff();
                            break;
                        case 310:
                            tariff_310_sum+=sortedData.get(i).getTarriff();
                            break;
                        case 270:
                            tariff_270_sum+=sortedData.get(i).getTarriff();
                            break;
                        default:
                            tariff_other=(int)sortedData.get(i).getTarriff();
                            commentary+=tariff_other+" "+sortedData.get(i).getFio_applicant()+"; \n";
                            break;

                    }

                }
                tariff_70_sum_allRegion+=tariff_70_sum; tariff_110_sum_allRegion+=tariff_110_sum; tariff_130_sum_allRegion+=tariff_130_sum;
                tariff_220_sum_allRegion+=tariff_220_sum;tariff_260_sum_allRegion+=tariff_260_sum;
                tariff_310_sum_allRegion+=tariff_310_sum; tariff_270_sum_allRegion+=tariff_270_sum;
                tariff_total_region+=Sum_tariff_mfc;

                String tariff_70_sum_str="";String tariff_110_sum_str="";String tariff_130_sum_str="";
                String tariff_220_sum_str="";String tariff_260_sum_str="";
                String tariff_310_sum_str="";String tariff_270_sum_str="";
                String tariff_other_sum_str="";
                String Sum_tariff_mfc_str= String.valueOf(Sum_tariff_mfc);
                if (tariff_70_sum!=0) tariff_70_sum_str= String.valueOf(tariff_70_sum);
                if (tariff_110_sum!=0) tariff_110_sum_str=String.valueOf(tariff_110_sum);
                if (tariff_130_sum!=0) tariff_130_sum_str=String.valueOf(tariff_130_sum);
                if (tariff_220_sum!=0) tariff_220_sum_str=String.valueOf(tariff_220_sum);
                if (tariff_260_sum!=0) tariff_260_sum_str=String.valueOf(tariff_260_sum);
                if (tariff_310_sum!=0) tariff_310_sum_str=String.valueOf(tariff_310_sum);
                if (tariff_270_sum!=0) tariff_270_sum_str=String.valueOf(tariff_270_sum);
                //if (tariff_other_sum!=0) tariff_other_sum_str=String.valueOf(tariff_other_sum);

                int name_mfc_rep=name_mfc+1;
                paymentsList.add(new PaymentsModel(name_mfc_rep,mfcWithoutDublicates.get(finalName_mfc), "",commentary,
                        "", tariff_70_sum_str, tariff_110_sum_str, tariff_130_sum_str, tariff_220_sum_str,
                        tariff_260_sum_str, tariff_310_sum_str, tariff_270_sum_str, Sum_tariff_mfc_str));
                System.out.println("SUM MFC: "+Sum_tariff_mfc);
            }

            for (PaymentsModel paymentsModel : paymentsList) {
                System.out.println(paymentsModel.getNumber_list() + " " + paymentsModel.getName_mfc()+ " "+  paymentsModel.getTariff_70_sum()+" "+
                        paymentsModel.getTariff_110_sum()+ " "+ paymentsModel.getTariff_130_sum()+" "+ paymentsModel.getTariff_220_sum()+" "
                        +paymentsModel.getTariff_260_sum()+" "+ paymentsModel.getTariff_310_sum()+" "+ paymentsModel.getTariff_270_sum()+
                        " " +paymentsModel.getTotal_sum());
            }

            // Itog
            paymentsList.add(new PaymentsModel(0,"ИТОГО:", "","",
                    "", String.valueOf(tariff_70_sum_allRegion), String.valueOf(tariff_110_sum_allRegion),
                    String.valueOf(tariff_130_sum_allRegion), String.valueOf(tariff_220_sum_allRegion),
                    String.valueOf(tariff_260_sum_allRegion), String.valueOf(tariff_310_sum_allRegion),
                    String.valueOf(tariff_270_sum_allRegion), String.valueOf(tariff_total_region)));


        }
        ArrayList<PaymentsModel> paymentsListFinal= new ArrayList<>(paymentsList) ;
        return paymentsListFinal;
    }

    public void SaveReportToFile(ArrayList<PaymentsModel> listPayment) throws IOException {
        //System.out.println(text_test);
        // Создаем экземпляр класса FileChooser
        FileChooser fileChooser = new FileChooser();

        String lastPathDirectory=getLastDirectory();
        if (!lastPathDirectory.equals("")){
            fileChooser.setInitialDirectory(new File(lastPathDirectory));
        }
        // Устанавливаем список расширений для файла
        fileChooser.setInitialFileName("report_payments");// Устанавливаем название для файла
        // Список расширений для Excel
        FileChooser.ExtensionFilter extFilterExcel = new FileChooser.ExtensionFilter("Excel file (*.xlsx)", "*.xlsx");
        // Добавляем список расширений
        fileChooser.getExtensionFilters().add(extFilterExcel);

        FXMLLoader loader = new FXMLLoader();
        loader.setLocation(getClass().getResource("/payments_mfc/view/app.fxml"));
        try {
            loader.load();
        } catch (IOException e) {
            e.printStackTrace();
        }
        Parent root = loader.getRoot();
        appController AppController = loader.getController();
        Stage stage = new Stage();
        // Показываем диалоговое окно для сохранения файла
        File file = fileChooser.showSaveDialog(stage);
        // Если не нажать кнопка "Отмена"
        if (fileChooser.getSelectedExtensionFilter() != null) {
            // Если выбрано расширение для Excel
            if (fileChooser.getSelectedExtensionFilter().getExtensions().toString().equals("[*.xlsx]")) {
                System.out.println("SELECTED XLSX");
                // Если файл не пустой
                if (file != null) {
                    // Сохраняем файл

                    // Окно, которое уведомляет о загрузке файла
                    ButtonType ok_but = new ButtonType("OK", ButtonBar.ButtonData.OK_DONE); // Создание кнопки "Открыть отчёт"
                    ButtonType cancel_but = new ButtonType("Cancel", ButtonBar.ButtonData.CANCEL_CLOSE); // Создание кнопки "Открыть папку с отчётом"
                    Alert alert = new Alert(Alert.AlertType.INFORMATION, "Test", ok_but, cancel_but);
                    alert.setTitle("Загрузка отчёта...");
                    alert.setHeaderText("Идёт загрузка отчёта, подождите...");
                    alert.setContentText("После загрузки появится уведомление!");
                    alert.show();
                    // Скрываем кнопки в окне, чтобы пользователь случайно не нажал их
                    Button okButton = (Button) alert.getDialogPane().lookupButton(ok_but);
                    Button cancelButton = (Button) alert.getDialogPane().lookupButton(cancel_but);
                    okButton.setVisible(false);
                    cancelButton.setVisible(false);


                    cancelButton.fire(); // Закрываем окно с загрузкой
                    // Получение директорий, вызов функции скачивания отчёта
                    ArrayList<String> pathFileAndDir=SaveFileExcel(listPayment, file);
                    // Получаем путь для файла
                    String pathToFile=pathFileAndDir.get(0);
                    String absolutePathToFile=pathFileAndDir.get(1);

                    System.out.println(pathToFile + " " + absolutePathToFile);
                    // Отображаем окно с выбором: Открыть отчёт или открыть папку с отчётом
                    ButtonType openReport = new ButtonType("Открыть отчёт", ButtonBar.ButtonData.OK_DONE); // Создание кнопки "Открыть отчёт"
                    ButtonType openDir = new ButtonType("Открыть папку с отчётом", ButtonBar.ButtonData.CANCEL_CLOSE); // Создание кнопки "Открыть папку с отчётом"
                    Alert alert2 = new Alert(Alert.AlertType.INFORMATION, "Test", openReport, openDir);
                    alert2.setTitle("Загрузка завершена!"); // Название предупреждения
                    alert2.setHeaderText("Отчёт загружен!"); // Текст предупреждения
                    alert2.setContentText("Отчёт доступен в папке: " + pathToFile);
                    // Вызов подтверждения элемента
                    alert2.showAndWait().ifPresent(rs -> {
                        if (rs == openReport) { // Если выбрали открыть отчёт
                            try {
                                Desktop.getDesktop().open(new File(absolutePathToFile));
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        } else if (rs == openDir) { // Если выбрали открыт папку
                            try {
                                Desktop.getDesktop().open(new File(pathToFile));
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        }
                    });

                }
            }
        }
    }

    // Функция для установки стилей Excel
    private static XSSFCellStyle createStyleForTitleNew(XSSFWorkbook workbook) {
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        return style;
    }

    private static XSSFCellStyle createStyleForTotalSum(XSSFWorkbook workbook) {
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        return style;
    }

    // Функция сохранения файла в Excel
    public static ArrayList<String> SaveFileExcel(ArrayList<PaymentsModel> dataReportList, File file) throws IOException {
        // Создание книги Excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        // Создание листа
        XSSFSheet sheet = workbook.createSheet("Отчёт по платежам");

        int rownum = 0;
        Cell cell;
        Row row;
        // Установка стилей
        XSSFCellStyle style = createStyleForTitleNew(workbook);
        XSSFCellStyle styleTotalSum = createStyleForTotalSum(workbook);
        row = sheet.createRow(rownum);

        // Создание столбца "Название организации"
        cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("№ п/п");
        cell.setCellStyle(style);

        cell = row.createCell(1, CellType.STRING);
        cell.setCellValue("Отделы \"Мои документы\" по территории");
        cell.setCellStyle(style);
        // Создание столбца "Номер обращения"
        cell = row.createCell(2, CellType.STRING);
        cell.setCellValue("Комментарий к уточнению");
        cell.setCellStyle(style);
        // Создание столбца "Наименование обращения"
        cell = row.createCell(3, CellType.STRING);
        cell.setCellValue("Уточнение платежа");
        cell.setCellStyle(style);
        // Создание столбца "Дата создания"
        cell = row.createCell(4, CellType.STRING);
        cell.setCellValue("70");
        cell.setCellStyle(style);
        // Создание столбца "Дата окончания"
        cell = row.createCell(5, CellType.STRING);
        cell.setCellValue("110");
        cell.setCellStyle(style);
        // Создание столбца "Статус"
        cell = row.createCell(6, CellType.STRING);
        cell.setCellValue("130");
        cell.setCellStyle(style);
        // Создание столбца "Текущий шаг"
        cell = row.createCell(7, CellType.STRING);
        cell.setCellValue("220");
        cell.setCellStyle(style);
        // Создание столбца "Заявители"
        cell = row.createCell(8, CellType.STRING);
        cell.setCellValue("260");
        cell.setCellStyle(style);

        cell = row.createCell(9, CellType.STRING);
        cell.setCellValue("310");
        cell.setCellStyle(style);

        cell = row.createCell(10, CellType.STRING);
        cell.setCellValue("270");
        cell.setCellStyle(style);

        cell = row.createCell(11, CellType.STRING);
        cell.setCellValue("ИТОГО");
        cell.setCellStyle(style);




        // Перебор по данным
        for (PaymentsModel reportModel : dataReportList) {
            //System.out.println(mfc_model.getIdMfc() +"\t" +mfc_model.getNameMfc());
            rownum++;
            row = sheet.createRow(rownum);

            // Запись данных в отчёт
            cell = row.createCell(0, CellType.STRING);
            cell.setCellValue(reportModel.getNumber_list());

            cell = row.createCell(1, CellType.STRING);
            cell.setCellValue(reportModel.getName_mfc());

            cell = row.createCell(2, CellType.STRING);
            cell.setCellValue(reportModel.getCommentary());

            cell = row.createCell(3, CellType.STRING);
            cell.setCellValue(reportModel.getCommentary_clarification());

            cell = row.createCell(4, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_70_sum());

            cell = row.createCell(5, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_110_sum());

            cell = row.createCell(6, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_130_sum());

            cell = row.createCell(7, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_220_sum());

            cell = row.createCell(8, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_260_sum());

            cell = row.createCell(9, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_310_sum());

            cell = row.createCell(10, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_270_sum());

            cell = row.createCell(11, CellType.STRING);
            cell.setCellValue(reportModel.getTotal_sum());

            if (rownum==dataReportList.size()){
                // Запись данных в отчёт
                cell = row.createCell(0, CellType.STRING);
                cell.setCellValue(reportModel.getNumber_list());
                cell.setCellStyle(styleTotalSum);

                cell = row.createCell(1, CellType.STRING);
                cell.setCellValue(reportModel.getName_mfc());
                cell.setCellStyle(styleTotalSum);

                cell = row.createCell(2, CellType.STRING);
                cell.setCellValue(reportModel.getCommentary());
                cell.setCellStyle(styleTotalSum);

                cell = row.createCell(3, CellType.STRING);
                cell.setCellValue(reportModel.getCommentary_clarification());
                cell.setCellStyle(styleTotalSum);

                cell = row.createCell(4, CellType.STRING);
                cell.setCellValue(reportModel.getTariff_70_sum());
                cell.setCellStyle(styleTotalSum);

                cell = row.createCell(5, CellType.STRING);
                cell.setCellValue(reportModel.getTariff_110_sum());
                cell.setCellStyle(styleTotalSum);

                cell = row.createCell(6, CellType.STRING);
                cell.setCellValue(reportModel.getTariff_130_sum());
                cell.setCellStyle(styleTotalSum);

                cell = row.createCell(7, CellType.STRING);
                cell.setCellValue(reportModel.getTariff_220_sum());
                cell.setCellStyle(styleTotalSum);

                cell = row.createCell(8, CellType.STRING);
                cell.setCellValue(reportModel.getTariff_260_sum());
                cell.setCellStyle(styleTotalSum);

                cell = row.createCell(9, CellType.STRING);
                cell.setCellValue(reportModel.getTariff_310_sum());
                cell.setCellStyle(styleTotalSum);

                cell = row.createCell(10, CellType.STRING);
                cell.setCellValue(reportModel.getTariff_270_sum());
                cell.setCellStyle(styleTotalSum);

                cell = row.createCell(11, CellType.STRING);
                cell.setCellValue(reportModel.getTotal_sum());
                cell.setCellStyle(styleTotalSum);
            }

        }
        // Выравнивание столбцов по ширине
        autoSizeColumns(workbook);
        // Создания потока сохранения файла
        FileOutputStream outFile = new FileOutputStream(file);
        // Запись файла
        workbook.write(outFile);
        // Закрытие потока записи
        outFile.close();

        checkDirectory(file.getParent());
        ArrayList<String> pathFileAndDir=new ArrayList<String>();
        pathFileAndDir.add(file.getParent());
        pathFileAndDir.add(file.getAbsolutePath());

        return pathFileAndDir; // Возвращение пути и абсолютного пути файла

    }

    // Функция для автоматического выравнивания столбцов по ширине содержимого
    public static void autoSizeColumns(Workbook workbook) {
        int numberOfSheets = workbook.getNumberOfSheets(); // Получаем кол-во листов
        for (int i = 0; i < numberOfSheets; i++) { // Идём по кол-ву листов
            Sheet sheet = workbook.getSheetAt(i); // Получаем лист
            if (sheet.getPhysicalNumberOfRows() > 0) { // Если столбцов больше 0
                Row row = sheet.getRow(sheet.getFirstRowNum()); // Получить первую строку
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) { // Идём по строкам и выравниваем по ширине
                    Cell cell = cellIterator.next();
                    int columnIndex = cell.getColumnIndex();
                    sheet.autoSizeColumn(columnIndex);
                }
            }
        }
    }

    public static void checkDirectory(String lastPathToFile){
        // Получаем логин, пароль, куки, флаг чекбокса
        SettingsModel settingsModel=new SettingsModel(lastPathToFile);
        settingsModel.setLastPathToFile(lastPathToFile);
        Gson gson = new Gson();
        File f = new File("C:\\payments_mfc");
        try{
            if(f.mkdir()) {
                System.out.println("Directory Created");
                String content=gson.toJson(settingsModel);
                File file=new File("C:\\payments_mfc\\settingsPaymentsMFC.json");
                Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file), StandardCharsets.UTF_8));
                out.write(content);
                out.close();

            } else {
                System.out.println("Directory is not created");
                SaveLastPathInfo(lastPathToFile);
            }
        } catch(Exception e){
            e.printStackTrace();
        }
    }

    // Функция для сохранения последнего пути файла
    public static void SaveLastPathInfo(String lastPathToFile) throws IOException {
        // Путь к файлу
        File fileJson = new File("C:\\payments_mfc\\settingsPaymentsMFC.json");
        // Проверяем, существует ли файл
        if(!fileJson.exists())
        {
            System.out.println("No file!");
        } else {
            System.out.println("yes file!");
            SettingsModel settingsModel=new SettingsModel(lastPathToFile);
            settingsModel.setLastPathToFile(lastPathToFile);

            // Сохраняем путь к файлу
            Gson gson = new Gson();
            File f = new File("C:\\payments_mfc");
            try{
                if(f.mkdir()) {
                    System.out.println("Directory Created");
                } else {
                    System.out.println("Directory is not created");
                }
            } catch(Exception e){
                e.printStackTrace();
            }

            String content=gson.toJson(settingsModel);
            // Записываем в кодировке UTF-8
            Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(fileJson), StandardCharsets.UTF_8));
            out.write(content);
            out.close();
        }
    }

    // Функция для получения последней директории, где был сохранён отчёт
    public String getLastDirectory(){
        String lastPathToFile="";
        // Читаем файл
        File fileJson = new File("C:\\payments_mfc\\settingsPaymentsMFC.json");

        if(!fileJson.exists())
        {
            System.out.println("No file!");
        } else {
            System.out.println("yes file!");
            // Парсим нужные данные
            JsonParser parser = new JsonParser();
            JsonElement jsontree = null;
            try {
                jsontree = parser.parse(new BufferedReader(new InputStreamReader(new FileInputStream("C:\\payments_mfc\\settingsPaymentsMFC.json"), StandardCharsets.UTF_8)));
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            // Получаем путь к файлу
            JsonObject jsonObject = jsontree.getAsJsonObject();
            lastPathToFile = jsonObject.get("lastPathToFile").getAsString();
        }
        return lastPathToFile;
    }


}
