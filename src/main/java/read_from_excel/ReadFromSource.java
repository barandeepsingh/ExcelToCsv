package read_from_excel;


import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ReadFromSource {
    public static void main(String[] args) {
//        String filePathExcel = "/Users/baran/Documents/linux_bootcamp/input_file.xlsx";
        String filePathCsv = "/Users/baran/Documents/linux_bootcamp/input_file.csv";
        String outputPath = "/Users/baran/Documents/linux_bootcamp/output_file.csv";
        String outputCsvHeader = "Col0,Col1,Col2,Col3,Col4,Col5";
        Map<Integer,Integer> mapColumns = new HashMap<>();
        mapColumns.put(1,3);
        mapColumns.put(0,0);
        mapColumns.put(2,4);
        mapColumns.put(5,5);
        int totalColumns = (int) Arrays.stream(outputCsvHeader.split(",")).count();
        String sheetName = "Sheet1";
//        List<String> dataFromExcel = FileUtils.getDataFromExcel(outputCsvHeader, filePathExcel, sheetName);
        List<String> dataFromCsv = FileUtils.getDataFromCsv(outputCsvHeader, filePathCsv, sheetName);
        FileUtils.writeToCsv(outputPath, dataFromCsv,mapColumns,totalColumns);
        System.out.println("Processing completed. File written at "+outputPath);

    }

}
