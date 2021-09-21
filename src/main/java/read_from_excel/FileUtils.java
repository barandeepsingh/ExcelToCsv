package read_from_excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class FileUtils {

    public static void writeToCsv(String outputFilePath, List<String> dataToWrite, Map<Integer,Integer> mapColumns, int totalColumns) {
        try {
            List<String> stringList = dataToWrite.stream().map(inputString -> changeMappingOfColumns(inputString, mapColumns, totalColumns)).collect(Collectors.toList());
            stringList.set(0,dataToWrite.get(0));
            Files.write(Paths.get(outputFilePath), stringList);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String changeMappingOfColumns(String inputString, Map<Integer,Integer> mapColumns, int totalColumns) {
        String[] outputArray = new String[totalColumns];
        Arrays.fill(outputArray,"");
        String[] inputArray = inputString.split(",");
        IntStream.rangeClosed(0,totalColumns).forEach(columnNumber->{
            if(mapColumns.containsKey(columnNumber)){
                int resultColumn = mapColumns.get(columnNumber);
                outputArray[resultColumn] = inputArray[columnNumber];
            }
        });
        return String.join(",",outputArray);
    }

    public static List<String> getDataFromExcel(String firstLine, String filePath, String sheetName) {
        ArrayList<String> finalList = new ArrayList<>();
        //Header for the final csv file
        finalList.add(firstLine);
        if (filePath != null && !"".equals(filePath.trim()) && sheetName != null && !"".equals(sheetName.trim())) {
            /* First need to open the file. */
            try (FileInputStream fInputStream = new FileInputStream(filePath.trim())) {

                /* Create the workbook object. */
                try (Workbook excelWorkBook = new XSSFWorkbook(fInputStream)) {

                    /* Get the sheet by name. */
                    Sheet employeeSheet = excelWorkBook.getSheet(sheetName);

                    int firstRowNum = employeeSheet.getFirstRowNum();
                    int lastRowNum = employeeSheet.getLastRowNum();

                    /* Because first row is header row, so we read data from second row. */
                    for (int i = firstRowNum + 1; i < lastRowNum + 1; i++) {
                        Row row = employeeSheet.getRow(i);

                        int firstCellNum = row.getFirstCellNum();
                        int lastCellNum = row.getLastCellNum();

                        StringJoiner stringJoiner = new StringJoiner(",");
                        for (int j = firstCellNum; j < lastCellNum; j++) {
                            stringJoiner.add(getValueOfCell(row, j));
                        }
                        finalList.add(stringJoiner.toString());
                    }
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return finalList;
    }

    private static String getValueOfCell(Row row, int cellNumber) {
        String cellValue = "";
        Cell cell = row.getCell(cellNumber);
        if (Objects.nonNull(cell)) {
            switch (cell.getCellType()) {
                case BOOLEAN:
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                case NUMERIC:
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                case STRING:
                    cellValue = cell.getStringCellValue();
                    break;
                case BLANK:
                case ERROR:
                    break;
            }
        }
        return cellValue;
    }

}
