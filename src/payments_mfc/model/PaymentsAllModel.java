package payments_mfc.model;

import java.util.ArrayList;

public class PaymentsAllModel {
    private int number_list;
    private double tarriff;
    private String name_mfc;
    private String commentary;
    private String commentary_clarification;
    private String serviceName;
    private double total_sum;

    public PaymentsAllModel(int number_list, double tarriff, String name_mfc, String commentary, String commentary_clarification, String serviceName, double total_sum) {
        this.number_list = number_list;
        this.tarriff = tarriff;
        this.name_mfc = name_mfc;
        this.commentary = commentary;
        this.commentary_clarification = commentary_clarification;
        this.serviceName = serviceName;
        this.total_sum = total_sum;
    }

    public int getNumber_list() {
        return number_list;
    }

    public void setNumber_list(int number_list) {
        this.number_list = number_list;
    }

    public double getTarriff() {
        return tarriff;
    }

    public void setTarriff(double tarriff) {
        this.tarriff = tarriff;
    }

    public String getName_mfc() {
        return name_mfc;
    }

    public void setName_mfc(String name_mfc) {
        this.name_mfc = name_mfc;
    }

    public String getCommentary() {
        return commentary;
    }

    public void setCommentary(String commentary) {
        this.commentary = commentary;
    }

    public String getCommentary_clarification() {
        return commentary_clarification;
    }

    public void setCommentary_clarification(String commentary_clarification) {
        this.commentary_clarification = commentary_clarification;
    }

    public String getServiceName() {
        return serviceName;
    }

    public void setServiceName(String serviceName) {
        this.serviceName = serviceName;
    }

    public double getTotal_sum() {
        return total_sum;
    }

    public void setTotal_sum(double total_sum) {
        this.total_sum = total_sum;
    }
}
