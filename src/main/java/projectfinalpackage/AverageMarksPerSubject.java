package projectfinalpackage;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class AverageMarksPerSubject {
    public static void averageMarksPerSubject(String excelFile,String wordFile) throws Exception {


        String excelFilePath = excelFile;
        ArrayList<Double> averageMarksPerSubjectList = new ArrayList<>();
        ArrayList<String> subjectList = new ArrayList<>();


        int subjectTotal = 0;
        int totalStudentCount = 0;
        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getLastRowNum();
            Row firstRow = sheet.getRow(0);
            int columnCount = firstRow.getLastCellNum();
            for (int j = 2; j < columnCount - 3; j++) {
            	 subjectTotal = 0;
                 totalStudentCount = 0;
                Cell cell = firstRow.getCell(j);
                if (cell != null) {
                    String subject = cell.getStringCellValue();
                    if ( subject.endsWith("[T]")) {
                        for (int i = 1; i <= rowCount; i++) {
                            Row row = sheet.getRow(i);
                            Cell gradeCell = row.getCell(j);
                            if (gradeCell != null) {
                                int gradeValue = convertGradeToValue(gradeCell.getStringCellValue());
                                subjectTotal += gradeValue;
                                totalStudentCount++;
                            }
                        }
                        averageMarksPerSubjectList.add(averageMarksPerSubject(subjectTotal,totalStudentCount));
                        subjectList.add(subject);


                    }
                }
            }

            String wordFilePathForAnalysis=wordFile;

            BarGraph.createGraph(wordFilePathForAnalysis,averageMarksPerSubjectList, subjectList,"Average Percentage Of Marks");

            workbook.close();
            inputStream.close();
        } catch (Exception e) {

        	javax.swing.JOptionPane.showMessageDialog(null, new javax.swing.JPanel(), e+"average marks class exception exception generated! ", javax.swing.JOptionPane.PLAIN_MESSAGE);

        }
    }

    private static int convertGradeToValue(String grade) {

    	if(grade.equalsIgnoreCase("A+")||grade.equalsIgnoreCase("A+*")||grade.equalsIgnoreCase("A+#")||grade.equalsIgnoreCase("A+**")||grade.equalsIgnoreCase("A+##")){
    		return 10;
    	}
    	if(grade.equalsIgnoreCase("A")||grade.equalsIgnoreCase("A*")||grade.equalsIgnoreCase("A#")||grade.equalsIgnoreCase("A**")||grade.equalsIgnoreCase("A##")){
    		return 9;
    	}
    	if(grade.equalsIgnoreCase("B+")||grade.equalsIgnoreCase("B+*")||grade.equalsIgnoreCase("B+#")||grade.equalsIgnoreCase("B+**")||grade.equalsIgnoreCase("B+##")){
    		return 8;
    	}
    	if(grade.equalsIgnoreCase("B")||grade.equalsIgnoreCase("B*")||grade.equalsIgnoreCase("B#")||grade.equalsIgnoreCase("B**")||grade.equalsIgnoreCase("B##")){
    		return 7;
    	}
    	if(grade.equalsIgnoreCase("C+")||grade.equalsIgnoreCase("C+*")||grade.equalsIgnoreCase("C+#")||grade.equalsIgnoreCase("C+**")||grade.equalsIgnoreCase("C+##")){
    		return 6;
    	}
    	if(grade.equalsIgnoreCase("C")||grade.equalsIgnoreCase("C*")||grade.equalsIgnoreCase("C#")||grade.equalsIgnoreCase("C**")||grade.equalsIgnoreCase("C##")){
    		return 5;
    	}
    	if(grade.equalsIgnoreCase("D+")||grade.equalsIgnoreCase("D+*")||grade.equalsIgnoreCase("D+#")||grade.equalsIgnoreCase("D+**")||grade.equalsIgnoreCase("D+##")){
    		return 4;
    	}
    	if(grade.equalsIgnoreCase("D")||grade.equalsIgnoreCase("D*")||grade.equalsIgnoreCase("D#")||grade.equalsIgnoreCase("D**")||grade.equalsIgnoreCase("D##")){
    		return 4;
    	}

    	return 0;

        }

    public static double averageMarksPerSubject(int totalMarks,int noOfStds) {

         double averageMarkPerSubject = ((double) totalMarks* 100 / (noOfStds*10) );
         return Double.parseDouble(String.format("%.2f", averageMarkPerSubject));
    }
    }

