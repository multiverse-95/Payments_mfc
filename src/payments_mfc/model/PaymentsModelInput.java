package payments_mfc.model;

public class PaymentsModelInput {
    private double tarriff;
    private String name_mfc;
    private String address_mfc;
    private String fio_applicant;
    private String service_mfc;

    public PaymentsModelInput(double tarriff, String name_mfc, String address_mfc, String fio_applicant, String service_mfc) {
        this.tarriff = tarriff;
        this.name_mfc = name_mfc;
        this.address_mfc = address_mfc;
        this.fio_applicant = fio_applicant;
        this.service_mfc = service_mfc;
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

    public String getAddress_mfc() {
        return address_mfc;
    }

    public void setAddress_mfc(String address_mfc) {
        this.address_mfc = address_mfc;
    }

    public String getFio_applicant() {
        return fio_applicant;
    }

    public void setFio_applicant(String fio_applicant) {
        this.fio_applicant = fio_applicant;
    }

    public String getService_mfc() {
        return service_mfc;
    }

    public void setService_mfc(String service_mfc) {
        this.service_mfc = service_mfc;
    }
}
