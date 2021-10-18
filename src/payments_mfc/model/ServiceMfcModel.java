package payments_mfc.model;

import java.util.ArrayList;

public class ServiceMfcModel {
    private int number_list;
    private ArrayList<PaymentsAllModel> listService;
    private double sumServices;

    public ServiceMfcModel(int number_list, ArrayList<PaymentsAllModel> listService, double sumServices) {
        this.number_list = number_list;
        this.listService = listService;
        this.sumServices = sumServices;
    }

    public int getNumber_list() {
        return number_list;
    }

    public void setNumber_list(int number_list) {
        this.number_list = number_list;
    }

    public ArrayList<PaymentsAllModel> getListService() {
        return listService;
    }

    public void setListService(ArrayList<PaymentsAllModel> listService) {
        this.listService = listService;
    }

    public double getSumServices() {
        return sumServices;
    }

    public void setSumServices(double sumServices) {
        this.sumServices = sumServices;
    }
}
