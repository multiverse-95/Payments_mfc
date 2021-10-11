package payments_mfc.model;

public class SettingsModel {
    private String lastPathToFile; // Последний путь к файлу, где был сохранен отчёт

    public SettingsModel(String lastPathToFile) {
        this.lastPathToFile = lastPathToFile;
    }

    public String getLastPathToFile() {
        return lastPathToFile;
    }

    public void setLastPathToFile(String lastPathToFile) {
        this.lastPathToFile = lastPathToFile;
    }
}
