package payments_mfc.controller;

import javafx.concurrent.Task;
import payments_mfc.model.PaymentsModel;

import java.time.LocalDate;
import java.util.ArrayList;

public class PaymentsController {

    // Класс для потока получения отчёта
    public static class ReportPayments extends Task<ArrayList<PaymentsModel>> {
        private final LocalDate datePeriod; // Дата начала в юникс


        public ReportPayments(LocalDate datePeriod) {
           this.datePeriod =datePeriod;
        }
        @Override
        protected ArrayList<PaymentsModel> call() throws Exception {
            ArrayList<PaymentsModel> reportListFinal=new ArrayList<PaymentsModel>();
            // получение одного общего отчёта
            //ArrayList<PaymentsModel> reportListFinal= GetReportAll(cookies,dateStartStr,dateFinishStr, MFCsList);
            return reportListFinal; // Возвращаем итоговый отчёт
        }
    }


}
