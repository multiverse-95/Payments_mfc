package payments_mfc.model;

public class PaymentsModel {
    private int number_list;
    private String name_mfc;
    private String commentary;
    private String commentary_clarification;
    private String tarriff;
    private String tariff_70_sum;
    private String tariff_110_sum;
    private String tariff_130_sum;
    private String tariff_220_sum;
    private String tariff_260_sum;
    private String tariff_310_sum;
    private String tariff_270_sum;
    private String total_sum;

    public PaymentsModel(int number_list, String name_mfc, String commentary, String commentary_clarification, String tarriff,
                         String tariff_70_sum, String tariff_110_sum, String tariff_130_sum, String tariff_220_sum, String tariff_260_sum,
                         String tariff_310_sum, String tariff_270_sum, String total_sum) {
        this.number_list = number_list;
        this.name_mfc = name_mfc;
        this.commentary = commentary;
        this.commentary_clarification = commentary_clarification;
        this.tarriff = tarriff;
        this.tariff_70_sum = tariff_70_sum;
        this.tariff_110_sum = tariff_110_sum;
        this.tariff_130_sum = tariff_130_sum;
        this.tariff_220_sum = tariff_220_sum;
        this.tariff_260_sum = tariff_260_sum;
        this.tariff_310_sum = tariff_310_sum;
        this.tariff_270_sum = tariff_270_sum;
        this.total_sum = total_sum;
    }

    public int getNumber_list() {
        return number_list;
    }

    public void setNumber_list(int number_list) {
        this.number_list = number_list;
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

    public String getTarriff() {
        return tarriff;
    }

    public void setTarriff(String tarriff) {
        this.tarriff = tarriff;
    }

    public String getTariff_70_sum() {
        return tariff_70_sum;
    }

    public void setTariff_70_sum(String tariff_70_sum) {
        this.tariff_70_sum = tariff_70_sum;
    }

    public String getTariff_110_sum() {
        return tariff_110_sum;
    }

    public void setTariff_110_sum(String tariff_110_sum) {
        this.tariff_110_sum = tariff_110_sum;
    }

    public String getTariff_130_sum() {
        return tariff_130_sum;
    }

    public void setTariff_130_sum(String tariff_130_sum) {
        this.tariff_130_sum = tariff_130_sum;
    }

    public String getTariff_220_sum() {
        return tariff_220_sum;
    }

    public void setTariff_220_sum(String tariff_220_sum) {
        this.tariff_220_sum = tariff_220_sum;
    }

    public String getTariff_260_sum() {
        return tariff_260_sum;
    }

    public void setTariff_260_sum(String tariff_260_sum) {
        this.tariff_260_sum = tariff_260_sum;
    }

    public String getTariff_310_sum() {
        return tariff_310_sum;
    }

    public void setTariff_310_sum(String tariff_310_sum) {
        this.tariff_310_sum = tariff_310_sum;
    }

    public String getTariff_270_sum() {
        return tariff_270_sum;
    }

    public void setTariff_270_sum(String tariff_270_sum) {
        this.tariff_270_sum = tariff_270_sum;
    }

    public String getTotal_sum() {
        return total_sum;
    }

    public void setTotal_sum(String total_sum) {
        this.total_sum = total_sum;
    }
}
