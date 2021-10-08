package payments_mfc.model;

public class PaymentsModelInput {
    private double tarriff;
    private String name_mfc;
    private String fio_applicant;

    public PaymentsModelInput(double tarriff, String name_mfc, String fio_applicant) {
        this.tarriff = tarriff;
        this.name_mfc = name_mfc;
        this.fio_applicant = fio_applicant;
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

    public String getFio_applicant() {
        return fio_applicant;
    }

    public void setFio_applicant(String fio_applicant) {
        this.fio_applicant = fio_applicant;
    }
}
