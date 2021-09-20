package read_from_excel;


import java.util.List;

public class ReadFromExcel {
    public static void main(String[] args) {
        String filePath = "/Users/baran/Documents/linux_bootcamp/input_file.xlsx";
        String outputPath = "/Users/baran/Documents/linux_bootcamp/output_file.csv";
        String outputCsvHeader = "Name,Age,Gender,IsAnEmployee";
        String sheetName = "Sheet1";
        List<String> dataFromExcel = FileUtils.getDataFromExcel(outputCsvHeader, filePath, sheetName);
        FileUtils.writeToCsv(outputPath, dataFromExcel);
        System.out.println("Processing completed. File written at "+outputPath);

    }

}
