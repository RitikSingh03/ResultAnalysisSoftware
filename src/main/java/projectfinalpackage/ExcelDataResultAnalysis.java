package projectfinalpackage;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class ExcelDataResultAnalysis {

    public static void excelDataResultAnalysis(String excelFile,String wordFile) {

    	  String excelFilePath = excelFile; // Path to your Excel file


          String wordFilePath = wordFile; /// Path to save Word file
        int cgpaColumnIndex = -1; // Define column index for CGPA
        int resultColumnIndex = -1; // Define column index for Result

        double passPercentage = .2f;
        int totalStudentsAppearing = 0;
        int passCount = 0;
        int failCount=0;
        int passHonsCount = 0;
        int passDiv1Count = 0;
        int passDiv2Count = 0;

        try (FileInputStream inputStream = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);

            // Find the column indices for CGPA and Result
            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                String header = cell.getStringCellValue().trim();
                if (header.equalsIgnoreCase("CGPA")) {
                    cgpaColumnIndex = cell.getColumnIndex();
                } else if (header.equalsIgnoreCase("Result")) {
                    resultColumnIndex = cell.getColumnIndex();
                }
            }

            if (cgpaColumnIndex == -1 || resultColumnIndex == -1) {
                return;
            }

            // Create map to hold student names and CGPA
            Map<String, Double> studentCGPAMap = new HashMap<>();

            // Retrieve student names and CGPA from each row
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Cell nameCell = row.getCell(1); // Assuming student name is in the second column
                Cell cgpaCell = row.getCell(cgpaColumnIndex);

                // Handle cases where name or CGPA cell is null or empty
                if (nameCell == null || nameCell.getCellType() == CellType.BLANK ||
                        cgpaCell == null || cgpaCell.getCellType() == CellType.BLANK) {
                    continue; // Skip this row
                }

                // Extract name and CGPA values
                String name = nameCell.getStringCellValue();
                double cgpa = getNumericValue(cgpaCell);

                // Add name and CGPA to the map
                studentCGPAMap.put(name, cgpa);
            }

            // Sort students by CGPA in descending order
            List<Map.Entry<String, Double>> sortedStudents = new ArrayList<>(studentCGPAMap.entrySet());
            sortedStudents.sort(Map.Entry.comparingByValue(Comparator.reverseOrder()));

            // Create Word document and add top 5 students based on CGPA
            try (FileOutputStream outputStream = new FileOutputStream(wordFilePath)) {
                XWPFDocument document = new XWPFDocument();

                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText("Top Five Students\n");

                // Create table
                XWPFTable table = document.createTable(6, 3); // 6 rows, 3 columns
                table.setCellMargins(100, 100, 100, 100);

                // Set column widths
                table.getRow(0).getCell(0).setWidth("500");
                table.getRow(0).getCell(1).setWidth("2000");
                table.getRow(0).getCell(2).setWidth("1000");

                // Add headers to the table
                table.getRow(0).getCell(0).setText("Sr. No.");
                table.getRow(0).getCell(1).setText("Student Name");
                table.getRow(0).getCell(2).setText("CGPA");

                // Add top 5 students to the table
                for (int i = 0; i < Math.min(5, sortedStudents.size()); i++) {
                    Map.Entry<String, Double> entry = sortedStudents.get(i);
                    String name = entry.getKey();
                    double cgpa = entry.getValue();

                    XWPFTableRow row = table.getRow(i + 1);
                    row.getCell(0).setText(Integer.toString(i + 1));
                    row.getCell(1).setText(name);
                    row.getCell(2).setText(Double.toString(cgpa));
                }


                totalStudentsAppearing = sheet.getLastRowNum();
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    Cell resultCell = row.getCell(resultColumnIndex);
                    Cell cgpaCell = row.getCell(cgpaColumnIndex);

                    if (resultCell == null || cgpaCell == null ||
                            resultCell.getCellType() == CellType.BLANK || cgpaCell.getCellType() == CellType.BLANK) {
                        continue;
                    }

                    String result = resultCell.getStringCellValue().trim();
                    result=result.toLowerCase();
                    if (result.equalsIgnoreCase("Pass") || result.startsWith("pass")) {
                        passCount++;
                        double cgpa = getNumericValue(cgpaCell);
                        if (cgpa >= 7.5) {
                            passHonsCount++;
                        }  if (cgpa >= 6.5 && cgpa<7.5) {
                            passDiv1Count++;
                        } if (cgpa >= 5.0 && cgpa < 6.5) {
                            passDiv2Count++;
                        }
                    }
                    else if(result.equalsIgnoreCase("Fail")){
                    	failCount++;
                    }
                }
                passPercentage = (double) passCount / totalStudentsAppearing * 100;
                XWPFParagraph statsParagraph = document.createParagraph();
                XWPFRun statsRun = statsParagraph.createRun();
                statsRun.setText("Pass percentage: " + String.format("%.2f", passPercentage));
                statsRun.addBreak();
                statsRun.setText("Total Students Appearing: " + totalStudentsAppearing);
                statsRun.addBreak();
                statsRun.setText("No. of students pass: " + passCount);
                statsRun.addBreak();
                statsRun.setText("No. of students fail: " + failCount);
                statsRun.addBreak();
                statsRun.setText("No. of students passed with Hons.: " + passHonsCount);
                statsRun.addBreak();
                statsRun.setText("No. of students passed in I Div.: " + passDiv1Count);
                statsRun.addBreak();
                statsRun.setText("No. of students passed in II Div.: " + passDiv2Count);

                document.write(outputStream);
            }



        } catch (Exception e) {
        	javax.swing.JOptionPane.showMessageDialog(null, new javax.swing.JPanel(), e+"exceldataresultanalysis class exception exception generated! ", javax.swing.JOptionPane.PLAIN_MESSAGE);

        }
    }

    private static double getNumericValue(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(cell.getStringCellValue());
            } catch (NumberFormatException e) {
                return 0;
            }
        }
        return 0;
    }
}



