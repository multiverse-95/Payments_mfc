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
import javafx.geometry.Pos;
import javafx.scene.Cursor;
import javafx.scene.Node;
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
import javafx.scene.text.Font;
import javafx.scene.text.FontPosture;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;


import payments_mfc.controller.PaymentsAllController;
import payments_mfc.controller.PaymentsController;
import payments_mfc.model.PaymentsItogModel;
import payments_mfc.model.PaymentsModel;

import java.io.File;
import java.io.FileInputStream;
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
    private StackPane root_all_report;

    @FXML
    private VBox vbox_rep_all_main;

    @FXML
    private DatePicker date_start_org_d;

    @FXML
    private DatePicker date_finish_org_d;

    @FXML
    private Button choose_file_all_b;

    @FXML
    private Button download_report_all_b;

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
        // Установить курсор при наведении на кнопки
        choose_file_b.setOnMouseEntered(event_mouse -> {
            ((Node) event_mouse.getSource()).setCursor(Cursor.HAND);
        });
        download_report_b.setOnMouseEntered(event_mouse -> {
            ((Node) event_mouse.getSource()).setCursor(Cursor.HAND);
        });
        choose_file_all_b.setOnMouseEntered(event_mouse -> {
            ((Node) event_mouse.getSource()).setCursor(Cursor.HAND);
        });
        download_report_all_b.setOnMouseEntered(event_mouse -> {
            ((Node) event_mouse.getSource()).setCursor(Cursor.HAND);
        });
        Label label_onStart=new Label();
        label_onStart.setText("Выберите файл с платежами, чтобы составить отчёт.");
        label_onStart.setFont(Font.font("verdana", FontWeight.BOLD, FontPosture.REGULAR, 15));
        data_rep_table.setPlaceholder(label_onStart);

        choose_file_b.setOnAction(event -> {
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
            PaymentsController paymentsController = new PaymentsController();
            String lastPathDirectory= paymentsController.getLastDirectory();
            if (!lastPathDirectory.equals("")){
                fileChooser.setInitialDirectory(new File(lastPathDirectory));
            }
            fileChooser.setTitle("Выберите файл с отчётом по платежам");//Заголовок диалога
            FileChooser.ExtensionFilter extFilter =
                    new FileChooser.ExtensionFilter("Excel file (*.xlsx)", "*.xlsx");//Расширение
            fileChooser.getExtensionFilters().add(extFilter);
            File file = fileChooser.showOpenDialog(stage);//Указываем текущую сцену CodeNote.mainStage
            paymentsController.checkDirectory(file.getParent());
            try {
                openCompareReport(file);
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        choose_file_all_b.setOnAction(event -> {
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
            PaymentsAllController paymentsAllController = new PaymentsAllController();
            String lastPathDirectory= paymentsAllController.getLastDirectory();
            if (!lastPathDirectory.equals("")){
                fileChooser.setInitialDirectory(new File(lastPathDirectory));
            }
            fileChooser.setTitle("Выберите файл с отчётом по платежам (все)");//Заголовок диалога
            FileChooser.ExtensionFilter extFilter =
                    new FileChooser.ExtensionFilter("Excel file (*.xlsx)", "*.xlsx");//Расширение
            fileChooser.getExtensionFilters().add(extFilter);
            File file = fileChooser.showOpenDialog(stage);//Указываем текущую сцену CodeNote.mainStage
            paymentsAllController.checkDirectory(file.getParent());
            try {
                openCompareReportAllPayments(file);
            } catch (IOException e) {
                e.printStackTrace();
            }

        });

    }

    public void openCompareReport(File file) throws IOException {

        ProgressIndicator pi = new ProgressIndicator(); // Запуск прогресс индикатора
        VBox box = new VBox(pi);
        box.setAlignment(Pos.CENTER);
        root_report.getChildren().add(box);
        vbox_rep_main.setDisable(true);
        data_rep_table.setDisable(true);

        // Запустить поток для получения отчёта с сервера
        Task ReportPaymentsTask = new PaymentsController.ReportPayments(file);

        //  После выполнения потока
        ReportPaymentsTask.setOnSucceeded(new EventHandler<WorkerStateEvent>() {
            @Override
            public void handle(WorkerStateEvent event) {
                box.setDisable(true);
                pi.setVisible(false);
                vbox_rep_main.setDisable(false);
                data_rep_table.setDisable(false);
                // Получение данных с распарсенного поля
                ArrayList<PaymentsModel> paymentsList= (ArrayList<PaymentsModel>) ReportPaymentsTask.getValue();

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
                download_report_b.setDisable(false);

                download_report_b.setOnAction(eventDownload -> {
                    try {
                        PaymentsController paymentsController=new PaymentsController();
                        paymentsController.SaveReportToFile(paymentsList);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                });
            }
        });
        // Запуск потока
        Thread reportThread = new Thread(ReportPaymentsTask);
        reportThread.setDaemon(true);
        reportThread.start();

    }

    public void openCompareReportAllPayments(File file) throws IOException {

        ProgressIndicator pi = new ProgressIndicator(); // Запуск прогресс индикатора
        VBox box = new VBox(pi);
        box.setAlignment(Pos.CENTER);
        root_all_report.getChildren().add(box);
        vbox_rep_all_main.setDisable(true);


        // Запустить поток для получения отчёта с сервера
        Task ReportAllPaymentsTask = new PaymentsAllController.ReportPaymentsAll(file);

        //  После выполнения потока
        ReportAllPaymentsTask.setOnSucceeded(new EventHandler<WorkerStateEvent>() {
            @Override
            public void handle(WorkerStateEvent event) {
                box.setDisable(true);
                pi.setVisible(false);
                vbox_rep_all_main.setDisable(false);

                // Получение данных с распарсенного поля
                ArrayList<PaymentsItogModel> paymentsList= (ArrayList<PaymentsItogModel>) ReportAllPaymentsTask.getValue();


                download_report_all_b.setDisable(false);

                download_report_all_b.setOnAction(eventDownload -> {
                    try {
                        PaymentsAllController paymentsAllController=new PaymentsAllController();
                        paymentsAllController.SaveReportToFile(paymentsList);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                });
            }
        });
        // Запуск потока
        Thread reportAllThread = new Thread(ReportAllPaymentsTask);
        reportAllThread.setDaemon(true);
        reportAllThread.start();

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

}
