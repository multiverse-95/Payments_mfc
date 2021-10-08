package payments_mfc;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.collections.transformation.SortedList;
import javafx.concurrent.Task;
import javafx.concurrent.WorkerStateEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.MenuBar;
import javafx.scene.control.MenuItem;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.HBox;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;


import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import payments_mfc.model.PaymentsModel;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.*;
import java.util.List;
import java.util.function.Predicate;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import payments_mfc.model.PaymentsModelInput;

public class appController {
    @FXML
    private ResourceBundle resources;

    @FXML
    private URL location;

    @FXML
    private MenuBar app_menuBar;

    @FXML
    private MenuItem menu_item_change_user;

    @FXML
    private StackPane root_report;

    @FXML
    private VBox vbox_rep_main;

    @FXML
    private DatePicker date_start_d;

    @FXML
    private DatePicker date_finish_d;

    @FXML
    private Button choose_file_b;

    @FXML
    private Button download_report_b;

    @FXML
    private HBox vbox_filter;

    @FXML
    private ChoiceBox<?> choiceFilter_box;

    @FXML
    private TextField search_t;

    @FXML
    private Button show_rep_b;

    @FXML
    private Label period_label;

    @FXML
    private TableView<PaymentsModel> data_rep_table;

    @FXML
    private TableColumn<PaymentsModel, Integer> number_list_col;

    @FXML
    private TableColumn<PaymentsModel, String> name_mfc_col;

    @FXML
    private TableColumn<PaymentsModel, String> commentary_col;

    @FXML
    private TableColumn<PaymentsModel, String> commentary_clarification_col;

    @FXML
    private TableColumn<PaymentsModel, String> tariff_70_sum_col;

    @FXML
    private TableColumn<PaymentsModel, String> tariff_110_sum_col;

    @FXML
    private TableColumn<PaymentsModel, String> tariff_130_sum_col;

    @FXML
    private TableColumn<PaymentsModel, String> tariff_220_sum_col;

    @FXML
    private TableColumn<PaymentsModel, String> tariff_260_sum_col;

    @FXML
    private TableColumn<PaymentsModel, String> tariff_310_sum_col;

    @FXML
    private TableColumn<PaymentsModel, String> tariff_270_sum_col;

    @FXML
    private TableColumn<PaymentsModel, String> total_sum_col;

    @FXML
    private StackPane root_org_report;

    @FXML
    private VBox vbox_rep_org_main;

    @FXML
    private DatePicker date_start_org_d;

    @FXML
    private DatePicker date_finish_org_d;

    @FXML
    private Button generate_report_org_b;

    @FXML
    private Button download_report_org_b;

    @FXML
    private HBox vbox_org_filter;

    @FXML
    private ChoiceBox<?> choiceFilter_org_box;

    @FXML
    private TextField search_org_t;

    @FXML
    private Button show_rep_org_b;

    @FXML
    private Label period_org_label;

    @FXML
    private TableView<?> data_rep_org_table;

    @FXML
    private TableColumn<?, ?> name_company_org_col;

    @FXML
    private TableColumn<?, ?> number_appeal_org_col;

    @FXML
    private TableColumn<?, ?> name_appeal_org_col;

    @FXML
    private TableColumn<?, ?> date_create_org_col;

    @FXML
    private TableColumn<?, ?> date_end_org_col;

    @FXML
    private TableColumn<?, ?> status_org_col;

    @FXML
    private TableColumn<?, ?> cur_step_org_col;

    @FXML
    private TableColumn<?, ?> applicant_org_col;

    @FXML
    void initialize() {
        final List<PaymentsModel>[] listPayments = new List[]{new ArrayList<PaymentsModel>()};
        choose_file_b.setOnAction(event -> {
            try {
                listPayments[0] = openCompareReport();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });
        download_report_b.setOnAction(event -> {
            try {
                SaveReportToFile(listPayments[0]);
            } catch (IOException e) {
                e.printStackTrace();
            }
        });
    }

    public List<PaymentsModel> openCompareReport() throws IOException {
        List<PaymentsModel> paymentsList = new ArrayList<PaymentsModel>();

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
        FileChooser fileChooser = new FileChooser();//Класс работы с диалогом выборки и сохранения
        // SaveLastPathController saveLastPathController=new SaveLastPathController();
        String lastPathDirectory= "D:\\recovery\\payment_reports\\testing_report";
        //lastPathDirectory="";
        if (!lastPathDirectory.equals("")){
            fileChooser.setInitialDirectory(new File(lastPathDirectory));
        }
        fileChooser.setTitle("Выберите файл с отчётом по платежам");//Заголовок диалога
        FileChooser.ExtensionFilter extFilter =
                new FileChooser.ExtensionFilter("Excel file (*.xlsx)", "*.xlsx");//Расширение
        fileChooser.getExtensionFilters().add(extFilter);
        File file = fileChooser.showOpenDialog(stage);//Указываем текущую сцену CodeNote.mainStage
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
                            commentary+=tariff_other+" "+sortedData.get(i).getFio_applicant()+" \n";
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


            ObservableList<PaymentsModel> dataReport = FXCollections.observableArrayList(paymentsList);
            // Заполнение данными таблицы
            number_list_col.setCellValueFactory(new PropertyValueFactory<>("number_list"));
            //number_list_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            name_mfc_col.setCellValueFactory(new PropertyValueFactory<>("name_mfc"));
            name_mfc_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            commentary_col.setCellValueFactory(new PropertyValueFactory<>("commentary"));
            commentary_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            commentary_clarification_col.setCellValueFactory(new PropertyValueFactory<>("commentary_clarification"));
            commentary_clarification_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            tariff_70_sum_col.setCellValueFactory(new PropertyValueFactory<>("tariff_70_sum"));
            tariff_70_sum_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            tariff_110_sum_col.setCellValueFactory(new PropertyValueFactory<>("tariff_110_sum"));
            tariff_110_sum_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            tariff_130_sum_col.setCellValueFactory(new PropertyValueFactory<>("tariff_130_sum"));
            tariff_130_sum_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            tariff_220_sum_col.setCellValueFactory(new PropertyValueFactory<>("tariff_220_sum"));
            tariff_220_sum_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            tariff_260_sum_col.setCellValueFactory(new PropertyValueFactory<>("tariff_260_sum"));
            tariff_220_sum_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            tariff_310_sum_col.setCellValueFactory(new PropertyValueFactory<>("tariff_310_sum"));
            tariff_310_sum_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            tariff_270_sum_col.setCellValueFactory(new PropertyValueFactory<>("tariff_270_sum"));
            tariff_270_sum_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            total_sum_col.setCellValueFactory(new PropertyValueFactory<>("total_sum"));
            total_sum_col.setCellFactory(TextFieldTableCell.<PaymentsModel>forTableColumn());

            data_rep_table.setItems(dataReport);
            autoResizeColumns(data_rep_table); // Выровнять колонки в таблице

        }
        download_report_b.setDisable(false);
        return paymentsList;
    }

    public void SaveReportToFile(List<PaymentsModel> listPayment) throws IOException {
        //System.out.println(text_test);
        // Создаем экземпляр класса FileChooser
        FileChooser fileChooser = new FileChooser();

        String lastPathDirectory = "D:\\recovery\\payment_reports\\testing_report";
        //lastPathDirectory="";
        if (!lastPathDirectory.equals("")) {
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
    public static ArrayList<String> SaveFileExcel(List<PaymentsModel> dataReportList, File file) throws IOException {
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


        ArrayList<String> pathFileAndDir=new ArrayList<String>();
        pathFileAndDir.add(file.getParent());
        pathFileAndDir.add(file.getAbsolutePath());

        return pathFileAndDir; // Возвращение пути и абсолютного пути файла

    }

    // Функция для выравнивания колонок в таблице по ширине текста
    public static void autoResizeColumns( TableView<PaymentsModel> table )
    {
        table.setColumnResizePolicy( TableView.UNCONSTRAINED_RESIZE_POLICY);
        table.getColumns().stream().forEach( (column) ->
        {
            // Получение минимальной ширины
            Text t = new Text( column.getText() );
            double max = t.getLayoutBounds().getWidth();
            for ( int i = 0; i < table.getItems().size(); i++ )
            {
                // Столбцы не должны быть пустыми
                if ( column.getCellData( i ) != null )
                {
                    t = new Text( column.getCellData( i ).toString() ); // Получить текст со столбца
                    double calcwidth = t.getLayoutBounds().getWidth(); // Получить ширину текста
                    // Запомнить новую макс ширину
                    if ( calcwidth > max )
                    {
                        max = calcwidth;
                    }
                }
            }
            // Добавить к максимальной ширине немного пространства
            column.setPrefWidth( max + 12.0d );
        } );
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


}
