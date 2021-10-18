package payments_mfc.model;

import java.util.ArrayList;

public class ServiceItogModel {
    private String serviceName;
    private ArrayList<PaymentsAllModel> listOfServices;
    private double serviceSum;

    public ServiceItogModel(String serviceName, ArrayList<PaymentsAllModel> listOfServices, double serviceSum) {
        this.serviceName = serviceName;
        this.listOfServices = listOfServices;
        this.serviceSum = serviceSum;
    }

    public String getServiceName() {
        return serviceName;
    }

    public void setServiceName(String serviceName) {
        this.serviceName = serviceName;
    }

    public ArrayList<PaymentsAllModel> getListOfServices() {
        return listOfServices;
    }

    public void setListOfServices(ArrayList<PaymentsAllModel> listOfServices) {
        this.listOfServices = listOfServices;
    }

    public double getServiceSum() {
        return serviceSum;
    }

    public void setServiceSum(double serviceSum) {
        this.serviceSum = serviceSum;
    }
}
