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
import javafx.concurrent.WorkerStateEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import payments_mfc.appController;
import payments_mfc.model.*;

import java.awt.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.List;
import java.util.function.Predicate;

public class PaymentsAllController {

    // Класс для потока получения отчёта
    public static class ReportPaymentsAll extends Task<ArrayList<PaymentsItogModel>> {
        //private final LocalDate datePeriod; // Дата начала в юникс
        private final File file;

        public ReportPaymentsAll(File file) {
            this.file = file;
        }
        @Override
        protected ArrayList<PaymentsItogModel> call() throws Exception {
            // получение одного общего отчёта
            ArrayList<PaymentsItogModel> reportListFinal;
            reportListFinal=GetReportAll(file);
            // ArrayList<PaymentsModel> reportListFinal= GetReportAll(file);
            return reportListFinal; // Возвращаем итоговый отчёт
        }
    }

    // Класс для потока скачивания отчёта
    public static class DownloadTaskExcel extends Task<ArrayList<String>> {
        private final ArrayList<PaymentsItogModel> dataReportList;
        private File file;

        public DownloadTaskExcel(ArrayList<PaymentsItogModel> dataReportList, File file) {
            this.dataReportList = dataReportList;
            this.file=file;

        }
        @Override
        protected ArrayList<String> call() throws Exception {
            // Получение файла Excel
            ArrayList<String> pathFileAndDir=SaveFileExcel(dataReportList, file);
            return pathFileAndDir; // Возвращаем путь к файлу
        }
    }

    public static ArrayList<PaymentsItogModel> GetReportAll(File file) throws IOException {

        List<PaymentsModel> paymentsList = new ArrayList<PaymentsModel>();
        ArrayList<PaymentsItogModel> itogPayments= new ArrayList<PaymentsItogModel>();
        ArrayList<ServiceItogModel> itogService= new ArrayList<ServiceItogModel>();
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
            List<PaymentsModelInput> paymentsListInputAddressMFC= new ArrayList<PaymentsModelInput>();
            List<PaymentsModelInput> paymentsListInputFioApplicant= new ArrayList<PaymentsModelInput>();
            List<PaymentsModelInput> paymentsModelInputService = new ArrayList<PaymentsModelInput>();
            List<PaymentsModelInput> paymentsListInputNameMFCAndTarriff= new ArrayList<PaymentsModelInput>();
            for (int cell_num=1; cell_num<=6; cell_num++){
                for (Row row : sheet) { // For each Row.
                    Cell cell = row.getCell(cell_num); // Get the Cell at the Index / Column you want.
                    CellType cellType = cell.getCellTypeEnum();
                    int number_of_list=1;
                    switch (cellType) {
                        case _NONE:
                            break;
                        case BOOLEAN:
                            break;
                        case BLANK:
                            break;
                        case FORMULA:
                            break;
                        case NUMERIC:
                            //System.out.print(cell.getNumericCellValue());
                            if (cell_num==1){
                                paymentsListInputTarrif.add(new PaymentsModelInput(cell.getNumericCellValue(),"","","",""));
                            }

                            break;
                        case STRING:
                            if (cell_num==1){
                                paymentsListInputTarrif.add(new PaymentsModelInput(0,"","","",""));
                            }
                            if (cell_num==2){
                                paymentsListInputNameMFC.add(new PaymentsModelInput(0,cell.getStringCellValue(),"","",""));
                            }
                            if (cell_num==3){
                                paymentsListInputAddressMFC.add(new PaymentsModelInput(0,"",cell.getStringCellValue(),"",""));
                            }
                            if (cell_num==4){
                                paymentsListInputFioApplicant.add(new PaymentsModelInput(0,"","",cell.getStringCellValue(),""));
                            }
                            if (cell_num==6){
                                paymentsModelInputService.add(new PaymentsModelInput(0,"","","",cell.getStringCellValue()));
                            }
                            break;
                        case ERROR:
                            break;
                    }
                }
            }
            System.out.println("Size mfc "+paymentsListInputNameMFC.size()+" Size tarrif "+ paymentsListInputTarrif.size() +" Size address "+ paymentsListInputAddressMFC.size()+
                    " Size service "+ paymentsModelInputService.size());
            /*for (int i=0; i<paymentsListInputNameMFC.size(); i++){
                int numb=i+1;
                System.out.println("numb: "+numb+" "+ paymentsListInputTarrif.get(i).getName_mfc() +" "+ paymentsListInputTarrif.get(i).getTarriff());
                System.out.println("numb: "+numb+" "+ paymentsListInputNameMFC.get(i).getName_mfc() +" "+ paymentsListInputNameMFC.get(i).getTarriff());
            }*/

            for (int i=0; i<paymentsListInputTarrif.size(); i++){
                if (paymentsListInputAddressMFC.get(i).getAddress_mfc().contains("Курако")) {
                    //paymentsListInputNameMFC.add(new PaymentsModelInput(0,"Новокузнецкий р-н",""));
                    paymentsListInputNameMFCAndTarriff.add(new PaymentsModelInput(paymentsListInputTarrif.get(i).getTarriff(),"Новокузнецкий р-н",
                            "", paymentsListInputFioApplicant.get(i).getFio_applicant(),paymentsModelInputService.get(i).getService_mfc()));
                }
                else if (paymentsListInputAddressMFC.get(i).getAddress_mfc().contains("Григорченкова")) {
                    //paymentsListInputNameMFC.add(new PaymentsModelInput(0,"Ленинск-кузнецкий р-н",""));
                    paymentsListInputNameMFCAndTarriff.add(new PaymentsModelInput(paymentsListInputTarrif.get(i).getTarriff(),"Ленинск-кузнецкий р-н",
                            "", paymentsListInputFioApplicant.get(i).getFio_applicant(),paymentsModelInputService.get(i).getService_mfc()));
                }
                else if (paymentsListInputAddressMFC.get(i).getAddress_mfc().contains("г Прокопьевск, пр-кт Гагарина")) {
                    //paymentsListInputNameMFC.add(new PaymentsModelInput(0,"Прокопьевский р-н",""));
                    paymentsListInputNameMFCAndTarriff.add(new PaymentsModelInput(paymentsListInputTarrif.get(i).getTarriff(),"Прокопьевский р-н",
                            "", paymentsListInputFioApplicant.get(i).getFio_applicant(),paymentsModelInputService.get(i).getService_mfc()));
                }
                else if (paymentsListInputAddressMFC.get(i).getAddress_mfc().contains("г Юрга, ул Машиностроителей")) {
                    //paymentsListInputNameMFC.add(new PaymentsModelInput(0,"Юргинский р-н",""));
                    paymentsListInputNameMFCAndTarriff.add(new PaymentsModelInput(paymentsListInputTarrif.get(i).getTarriff(),"Юргинский р-н",
                            "", paymentsListInputFioApplicant.get(i).getFio_applicant(),paymentsModelInputService.get(i).getService_mfc()));
                }
                else {
                    paymentsListInputNameMFCAndTarriff.add(new PaymentsModelInput(paymentsListInputTarrif.get(i).getTarriff(),paymentsListInputNameMFC.get(i).getName_mfc(),
                            "", paymentsListInputFioApplicant.get(i).getFio_applicant(),paymentsModelInputService.get(i).getService_mfc()));
                }

            }

            /*for (int i=0; i< paymentsListInputNameMFCAndTarriff.size(); i++){
                int numb=i+1;
                System.out.println("numb: "+numb+" "+ paymentsListInputNameMFCAndTarriff.get(i).getTarriff() +" "+ paymentsListInputNameMFCAndTarriff.get(i).getName_mfc());
            }*/
            paymentsListInputNameMFCAndTarriff.remove(0);
            /*for (int i=0; i< paymentsListInputNameMFCAndTarriff.size(); i++){
                int numb=i+1;
                System.out.println("numb: "+numb+" "+ paymentsListInputNameMFCAndTarriff.get(i).getTarriff() +" "+ paymentsListInputNameMFCAndTarriff.get(i).getName_mfc()+
                        " "+paymentsListInputNameMFCAndTarriff.get(i).getService_mfc());
            }*/
            List<String> mfcSwithDublicates=new ArrayList<>();
            List<String> servicesWithDublicates=new ArrayList<>();
            List<String> mfcWithoutDublicates= new ArrayList<>();
            List<String> servicesWithoutDublicates= new ArrayList<>();

            for (PaymentsModelInput modelInput : paymentsListInputNameMFCAndTarriff) {
                mfcSwithDublicates.add(modelInput.getName_mfc());
                servicesWithDublicates.add(modelInput.getService_mfc());
            }

            Set<String> setMfc = new HashSet<>(mfcSwithDublicates);
            Set<String> setServices = new HashSet<>(servicesWithDublicates);

            mfcWithoutDublicates.addAll(setMfc);
            servicesWithoutDublicates.addAll(setServices);

            // Get duplicates
            /*System.out.println("WITH DUBLICATES");
            for (int i=0; i< mfcSwithDublicates.size(); i++){
                int numb=i+1;
                //System.out.println("numb "+numb+ " MFC "+mfcSwithDublicates.get(i));
            }
            System.out.println("WITHOUT DUBLICATES MFCS");
            for (int i=0; i< mfcWithoutDublicates.size(); i++){
                int numb=i+1;
                //System.out.println("numb "+numb+ " MFC "+mfcWithoutDublicates.get(i));
            }

            System.out.println("WITHOUT DUBLICATES SERVICES");
            for (int i=0; i< servicesWithoutDublicates.size(); i++){
                int numb=i+1;
                //System.out.println("numb "+numb+ " MFC "+servicesWithoutDublicates.get(i));
            }*/
            //

            ArrayList<PaymentsAllModel> paymentsAllList= new ArrayList<PaymentsAllModel>();
            ArrayList<ServiceMfcModel> serviceMfcList = new ArrayList<ServiceMfcModel>();
            for (int name_mfc=0; name_mfc<mfcWithoutDublicates.size();name_mfc++ ){
                int Sum_tariff_mfc=0;
                ObservableList<PaymentsModelInput> data = FXCollections.observableArrayList(paymentsListInputNameMFCAndTarriff);
                FilteredList<PaymentsModelInput> filteredData = new FilteredList<>(data, e -> true);
                int finalName_mfc = name_mfc;
                filteredData.setPredicate((Predicate<? super PaymentsModelInput>) paymentsModelInput->{
                    String lowerCaseFilter = mfcWithoutDublicates.get(finalName_mfc).toLowerCase();
                    // Сравниваем обращения с фильтром
                    if (paymentsModelInput.getName_mfc().toLowerCase().contains(lowerCaseFilter)){
                        return true;
                    }
                    return false;
                });
                SortedList<PaymentsModelInput> sortedDataByMfc = new SortedList<>(filteredData);
                ArrayList<PaymentsModelInput> sortedDataByMfcLIST = new ArrayList<>(sortedDataByMfc);
                System.out.println("SORTED DATA!!! \n");
                for (int name_serv=0; name_serv<servicesWithoutDublicates.size(); name_serv++){
                    ObservableList<PaymentsModelInput> dataServices = FXCollections.observableArrayList(sortedDataByMfcLIST);
                    FilteredList<PaymentsModelInput> filteredDataByMFCServices = new FilteredList<>(dataServices, e -> true);
                    int final_Services = name_serv;
                    int finalName_serv = name_serv;
                    filteredDataByMFCServices.setPredicate((Predicate<? super PaymentsModelInput>) paymentsModelInput->{
                        String lowerCaseFilterServ = servicesWithoutDublicates.get(final_Services).toLowerCase();

                        // Сравниваем обращения с фильтром
                        if (paymentsModelInput.getService_mfc().toLowerCase().contains(lowerCaseFilterServ)){
                            return true;
                        }
                        return false;
                    });
                    SortedList<PaymentsModelInput> sortedListMFCServices = new SortedList<>(filteredDataByMFCServices);

                    if (sortedListMFCServices.size()!=0){
                        double sum_Serv=0;

                        for (PaymentsModelInput sortedListMFCService : sortedListMFCServices) {
                            double sumTarif_serv=0;
                            System.out.println("Tariff: " + sortedListMFCService.getTarriff() + " Mfc: " + sortedListMFCService.getName_mfc() +
                                    " Name Service: " + sortedListMFCService.getService_mfc());
                            sumTarif_serv=sortedListMFCService.getTarriff();
                            paymentsAllList.add(new PaymentsAllModel(0,sortedListMFCService.getTarriff(),sortedListMFCService.getName_mfc(), "",
                                    "",sortedListMFCService.getService_mfc(), sumTarif_serv));
                            sum_Serv+=sortedListMFCService.getTarriff();
                        }
                        ArrayList<PaymentsAllModel> paymentsOneMFCList= new ArrayList<PaymentsAllModel>();
                        paymentsOneMFCList=paymentsAllList;

                        serviceMfcList.add(new ServiceMfcModel(0,paymentsOneMFCList, sum_Serv));
                        paymentsAllList= new ArrayList<PaymentsAllModel>();
                    }


                }
            }


            for (int i=0; i<mfcWithoutDublicates.size(); i++){
                double total_sum_services_city=0;
                for (int j=0; j<serviceMfcList.size(); j++){
                    if (serviceMfcList.get(j).getListService().get(0).getName_mfc().equals(mfcWithoutDublicates.get(i))){
                        double serveOne_Sum=0;
                        for (int smfcLi=0; smfcLi<serviceMfcList.get(j).getListService().size(); smfcLi++) {
                            serveOne_Sum+= serviceMfcList.get(j).getListService().get(smfcLi).getTarriff();
                        }
                        itogService.add(new ServiceItogModel(serviceMfcList.get(j).getListService().get(0).getServiceName(), serviceMfcList.get(j).getListService(),serveOne_Sum));
                        total_sum_services_city+=serviceMfcList.get(j).getSumServices();
                    }
                }
                itogPayments.add(new PaymentsItogModel(0,0,mfcWithoutDublicates.get(i),"","",
                        itogService, total_sum_services_city));
                itogService= new ArrayList<ServiceItogModel>();


            }
            System.out.println(paymentsAllList.size());
            System.out.println(serviceMfcList.size());
            System.out.println(itogPayments.size());
            System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!!");
            for (int i=0; i<itogPayments.size(); i++){
                for (int j=0; j<itogPayments.get(i).getServiceList().size(); j++){
                    System.out.println("MFC: "+itogPayments.get(i).getName_mfc() +" Service: "+itogPayments.get(i).getServiceList().get(j).getServiceName() +
                            " Service sum: "+itogPayments.get(i).getServiceList().get(j).getServiceSum());
                }

            }
        }


        return itogPayments;
    }

    public void SaveReportToFile(ArrayList<PaymentsItogModel> listPayment) throws IOException {
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
                    Alert alert =new Alert(Alert.AlertType.INFORMATION , "Test", ok_but, cancel_but);
                    alert.setTitle("Загрузка отчёта...");
                    alert.setHeaderText("Идёт загрузка отчёта, подождите...");
                    alert.setContentText("После загрузки появится уведомление!");
                    alert.show();
                    // Скрываем кнопки в окне, чтобы пользователь случайно не нажал их
                    Button okButton =( Button ) alert.getDialogPane().lookupButton( ok_but );
                    Button cancelButton = ( Button ) alert.getDialogPane().lookupButton( cancel_but );
                    okButton.setVisible(false);
                    cancelButton.setVisible(false);

                    Task DownloadTaskExcel =  new DownloadTaskExcel(listPayment, file);

                    //  После выполнения потока
                    DownloadTaskExcel.setOnSucceeded(new EventHandler<WorkerStateEvent>() {
                        @Override
                        public void handle(WorkerStateEvent event) {

                            cancelButton.fire(); // Закрываем окно с загрузкой
                            // Получение директорий, вызов функции скачивания отчёта
                            ArrayList<String> pathFileAndDir= (ArrayList<String>) DownloadTaskExcel.getValue();
                            // Получаем путь для файла
                            String pathToFile=pathFileAndDir.get(0);
                            String absolutePathToFile=pathFileAndDir.get(1);

                            System.out.println(pathToFile +" "+absolutePathToFile);
                            // Отображаем окно с выбором: Открыть отчёт или открыть папку с отчётом
                            ButtonType openReport = new ButtonType("Открыть отчёт", ButtonBar.ButtonData.OK_DONE); // Создание кнопки "Открыть отчёт"
                            ButtonType openDir = new ButtonType("Открыть папку с отчётом", ButtonBar.ButtonData.CANCEL_CLOSE); // Создание кнопки "Открыть папку с отчётом"
                            Alert alert =new Alert(Alert.AlertType.INFORMATION , "Test", openReport, openDir);
                            alert.setTitle("Загрузка завершена!"); // Название предупреждения
                            alert.setHeaderText("Отчёт загружен!"); // Текст предупреждения
                            alert.setContentText("Отчёт доступен в папке: "+pathToFile);
                            // Вызов подтверждения элемента
                            alert.showAndWait().ifPresent(rs -> {
                                if (rs == openReport){ // Если выбрали открыть отчёт
                                    try {
                                        Desktop.getDesktop().open(new File(absolutePathToFile));
                                    } catch (IOException e) {
                                        e.printStackTrace();
                                    }
                                } else if (rs==openDir){ // Если выбрали открыт папку
                                    try {
                                        Desktop.getDesktop().open(new File(pathToFile));
                                    } catch (IOException e) {
                                        e.printStackTrace();
                                    }
                                }
                            });
                        }
                    });

                    // Запуск потока
                    Thread DownloadThread = new Thread(DownloadTaskExcel);
                    DownloadThread.setDaemon(true);
                    DownloadThread.start();
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
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }
    // Функция для установки стилей Excel
    private static XSSFCellStyle createStyleForCells(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static XSSFCellStyle createStyleForTotalSum(XSSFWorkbook workbook) {
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    // Функция сохранения файла в Excel
    public static ArrayList<String> SaveFileExcel(ArrayList<PaymentsItogModel> dataReportList, File file) throws IOException {
        // Создание книги Excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        // Создание листа
        XSSFSheet sheet = workbook.createSheet("Отчёт по платежам");

        int rownum = 0;
        Cell cell;
        Row row;
        // Установка стилей
        XSSFCellStyle style = createStyleForTitleNew(workbook);
        XSSFCellStyle styleCells = createStyleForCells(workbook);
        XSSFCellStyle styleTotalSum = createStyleForTotalSum(workbook);
        row = sheet.createRow(rownum);


        ArrayList<String> mfcSwithDublicates= new ArrayList<String>();
        ArrayList<String> mfcWithoutDublicates= new ArrayList<String>();
        ArrayList<String> servicesWithDublicates=new ArrayList<String>();
        ArrayList<String> servicesWithoutDublicates=new ArrayList<String>();

        for (int i=0; i<dataReportList.size(); i++){
            for (int j=0; j<dataReportList.get(i).getServiceList().size(); j++){
                System.out.println("MFC: "+dataReportList.get(i).getName_mfc() +" Service: "+dataReportList.get(i).getServiceList().get(j).getServiceName() +
                        " Service sum: "+dataReportList.get(i).getServiceList().get(j).getServiceSum());
                mfcSwithDublicates.add(dataReportList.get(i).getName_mfc());
                servicesWithDublicates.add(dataReportList.get(i).getServiceList().get(j).getServiceName());
            }
        }

        Set<String> setMfc = new HashSet<>(mfcSwithDublicates);
        mfcWithoutDublicates.addAll(setMfc);
        Set<String> setServices = new HashSet<>(servicesWithDublicates);
        servicesWithoutDublicates.addAll(setServices);

        System.out.println(mfcWithoutDublicates.size()+" "+servicesWithoutDublicates.size());

        // Создание столбца "Название организации"
        cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("№ п/п");
        cell.setCellStyle(style);

        cell = row.createCell(1, CellType.STRING);
        cell.setCellValue("Отделы \"Мои документы\" по территории");
        cell.setCellStyle(style);

        cell = row.createCell(2, CellType.STRING);
        cell.setCellValue("Комментарий к уточнению");
        cell.setCellStyle(style);


        cell = row.createCell(3, CellType.STRING);
        cell.setCellValue("Уточнение платежа");
        cell.setCellStyle(style);


        int countServ=servicesWithoutDublicates.size();
        for (int k=0;k<=servicesWithoutDublicates.size(); k++){
            if (k==countServ){
                cell = row.createCell(countServ+4, CellType.STRING);
                cell.setCellValue("ИТОГО");
                cell.setCellStyle(style);
            } else {
                cell = row.createCell(k+4, CellType.STRING);
                cell.setCellValue(servicesWithoutDublicates.get(k));
                cell.setCellStyle(style);
            }
        }


        for (int mfcName=0; mfcName<mfcWithoutDublicates.size(); mfcName++){
            rownum++;
            row = sheet.createRow(rownum);
            // Запись данных в отчёт
            cell = row.createCell(0, CellType.STRING);
            cell.setCellValue(0);
            cell.setCellStyle(styleCells);

            // Запись данных в отчёт
            cell = row.createCell(1, CellType.STRING);
            cell.setCellValue(dataReportList.get(mfcName).getName_mfc());
            cell.setCellStyle(styleCells);

            // Запись данных в отчёт
            cell = row.createCell(2, CellType.STRING);
            cell.setCellValue("");
            cell.setCellStyle(styleCells);

            // Запись данных в отчёт
            cell = row.createCell(3, CellType.STRING);
            cell.setCellValue("");
            cell.setCellStyle(styleCells);

            for (int k=0;k<=servicesWithoutDublicates.size(); k++){
                if (k==countServ){
                    cell = row.createCell(countServ+4, CellType.STRING);
                    cell.setCellValue(dataReportList.get(mfcName).getTotal_sum());
                    cell.setCellStyle(styleCells);
                } else {
                    cell = row.createCell(k+4, CellType.STRING);
                    for (int sL=0; sL<dataReportList.get(mfcName).getServiceList().size();sL++){
                        if (servicesWithoutDublicates.get(k).equals(dataReportList.get(mfcName).getServiceList().get(sL).getServiceName())){
                            cell.setCellValue(dataReportList.get(mfcName).getServiceList().get(sL).getServiceSum());
                            cell.setCellStyle(styleCells);
                        }

                    }

                }
            }


        }

        /*// Перебор по данным
        for (PaymentsItogModel reportModel : dataReportList) {
            //System.out.println(mfc_model.getIdMfc() +"\t" +mfc_model.getNameMfc());
            rownum++;
            row = sheet.createRow(rownum);

            // Запись данных в отчёт
            cell = row.createCell(0, CellType.STRING);
            cell.setCellValue(reportModel.getNumber_list());
            cell.setCellStyle(styleCells);

            cell = row.createCell(1, CellType.STRING);
            cell.setCellValue(reportModel.getName_mfc());
            cell.setCellStyle(styleCells);

            cell = row.createCell(2, CellType.STRING);
            cell.setCellValue(reportModel.getCommentary());
            cell.setCellStyle(styleCells);

            cell = row.createCell(3, CellType.STRING);
            cell.setCellValue(reportModel.getCommentary_clarification());
            cell.setCellStyle(styleCells);

            cell = row.createCell(4, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_70_sum());
            cell.setCellStyle(styleCells);

            cell = row.createCell(5, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_110_sum());
            cell.setCellStyle(styleCells);

            cell = row.createCell(6, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_130_sum());
            cell.setCellStyle(styleCells);

            cell = row.createCell(7, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_220_sum());
            cell.setCellStyle(styleCells);

            cell = row.createCell(8, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_260_sum());
            cell.setCellStyle(styleCells);

            cell = row.createCell(9, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_310_sum());
            cell.setCellStyle(styleCells);

            cell = row.createCell(10, CellType.STRING);
            cell.setCellValue(reportModel.getTariff_270_sum());
            cell.setCellStyle(styleCells);

            cell = row.createCell(11, CellType.STRING);
            cell.setCellValue(reportModel.getTotal_sum());
            cell.setCellStyle(styleCells);

            if (rownum==dataReportList.size()){
                // Запись данных в отчёт
                cell = row.createCell(0, CellType.STRING);
                cell.setCellValue("");
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

        }*/
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
