package projectfinalpackage;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PassPercentSubject {
    public static void passPercentSubject(String excelFile,String wordFile) {
        String excelFilePath =excelFile ;
        ArrayList<Double> passPercentPerSubjectList = new ArrayList<>();
        ArrayList<String> subjectList = new ArrayList<>();


        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getLastRowNum();
            Row firstRow = sheet.getRow(0);
            int columnCount = firstRow.getLastCellNum();
            for (int j = 2; j < columnCount - 3; j++) {
                Cell cell = firstRow.getCell(j);
                if (cell != null) {
                    String subject = cell.getStringCellValue();
                    if (subject.endsWith("[T]")) {

                    	subjectList.add(subject);

                        int subjectCount = traverseColumn(excelFilePath, subject);

                        if (subjectCount != -1) {
                            passPercentPerSubjectList.add(percentFinder(subjectCount, rowCount));
                        }
                    }
                }
            }

            String wordFilePathForAnalysis=wordFile;
            BarGraph.createGraph(wordFilePathForAnalysis,passPercentPerSubjectList, subjectList,"Pass Percentage Of Students");
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public static int traverseColumn(String excelFilePath, String columnName) {
        try (FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
             Workbook workbook = WorkbookFactory.create(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);

            int columnIndex = -1;
            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equals(columnName)) {
                    columnIndex = cell.getColumnIndex();
                    break;
                }
            }

            if (columnIndex == -1) {
                System.out.println("Column '" + columnName + "' not found.");
                return -1;
            } else {
                int count = 0;

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    Cell cell = row.getCell(columnIndex);
                    if (cell != null) {
                        String value = cell.getStringCellValue();

                        if (!value.startsWith("f") && !value.startsWith("F")&&!value.contains("ABS")&&!value.contains("abs")) {


                            count++;
                        }
                    }
                }
                return count;
            }
        } catch (Exception ex) {
        	javax.swing.JOptionPane.showMessageDialog(null, new javax.swing.JPanel(), ex+"Pass Percent Exception generated ", javax.swing.JOptionPane.PLAIN_MESSAGE);

        	return -1;
        }
    }


    public static double percentFinder(int passCount, int totalStudent) {
        int noOfStd = totalStudent ;
        double passPercentPerSubject = ((double) passCount* 100 / noOfStd) ;
        return Double.parseDouble(String.format("%.2f", passPercentPerSubject));
    }
}
