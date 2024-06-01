package projectfinalpackage;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MarksExtractor {



    public static void marksExtractInExcel() {

    	ArrayList<String> pdfFilePaths = new ArrayList<>();

    File [] selectedFiles=CertificatePDFSelection.selectPdf();
    for (File file : selectedFiles) {
    	pdfFilePaths.add(file.getAbsolutePath());
    }




        String excelFilePath = ExcelFilePathSelection.excelFilePathSelection();

        if (excelFilePath!=null&&excelFilePath.toLowerCase().endsWith("XLSX".toLowerCase())) {
            try {
                convertPDFsToExcelAndExtractData(pdfFilePaths, excelFilePath);
            } catch (IOException e) {
                e.printStackTrace();
            }


        } else {

        }

    }

    private static void convertPDFsToExcelAndExtractData(List<String> pdfFilePaths, String excelFilePath) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("CombinedSheet");

            int rowIndex = 0;
            boolean isFirstPDF = true; // Flag to track if it's the first PDF

            List<String> firstPdfHeaders = null;

            for (int pdfIndex = 0; pdfIndex < pdfFilePaths.size(); pdfIndex++) {
                String pdfFilePath = pdfFilePaths.get(pdfIndex);

                try (PDDocument document = PDDocument.load(new File(pdfFilePath))) {
                    PDFTextStripper pdfStripper = new PDFTextStripper();
                    String text = pdfStripper.getText(document);

                    // Add patterns for name, enrollment number, and SGPA
                    Pattern namePattern = Pattern.compile("Name[ :]+([^\\n]+)");
                    Matcher nameMatcher = namePattern.matcher(text);

                    Pattern enrollmentPattern = Pattern.compile("Enrollment No.[ :]+([^\\n]+)");
                    Matcher enrollmentMatcher = enrollmentPattern.matcher(text);

                    Pattern sgpaPattern = Pattern.compile("SGPA[\\s\\S]*?(\\d+(?:\\.\\d+)?)");
                    Matcher sgpaMatcher = sgpaPattern.matcher(text);

                    // Update the subjectPattern to handle both types of subjects
                    Pattern subjectPattern = Pattern.compile("([A-Z0-9-]+)(?: \\[([TP])\\])? (.+?) (\\d+|\\-) (\\d+|\\-) ([A-Z+-]+)");
                    Matcher subjectMatcher = subjectPattern.matcher(text);

//                    Matcher subjectMatcher2=  subjectPattern2.matcher(text);

                    // Create a list to store the transposed data
                    List<List<String>> transposedData = new ArrayList<>();

                    String enrollmentNo = ""; // Initialize enrollment number variable
                    if (enrollmentMatcher.find()) {
                        // Extract and store the enrollment number
                        enrollmentNo = enrollmentMatcher.group(1).trim();
                    }
                    // Add Enrollment No. to transposed data
                    addDataToList(transposedData, "Enrollment No.", enrollmentNo);

                    // Extract and add Name if found
                    if (nameMatcher.find()) {
                        // Extract only the name from the full string
                        String fullName = nameMatcher.group(1).trim();
                        // Extract only the part before Enrollment No. if it is present in the name
                        String nameValue = fullName.split("Enrollment No")[0].trim();
                        addDataToList(transposedData, "Name", nameValue);
                    }

                    // Extract and add SGPA if found and it's an integer or decimal


                    // Extract and add each subject information
                    while (subjectMatcher.find()) {
                        String subjectCode = subjectMatcher.group(1);
                        String theoryOrPractical = subjectMatcher.group(2); // Captures the letter inside "[T]" or "[P]"
                        String subjectName = subjectMatcher.group(3);
                        String totalCredits = subjectMatcher.group(4);
                        String earnedCredits = subjectMatcher.group(5);
                        String grade = subjectMatcher.group(6);


                        // Determine if it's a theory or practical subject
                        // Add Subject Name and Grade information to the transposed data
                        String subjectHeader = (theoryOrPractical != null) ? " [" + theoryOrPractical + "]" : "";
                        addDataToList(transposedData, subjectName + " " + subjectHeader, grade);
                    }


                    if (sgpaMatcher.find()) {
                        addDataToList(transposedData, "SGPA", sgpaMatcher.group(1).trim());
                    }

                    // Add Result column
                    String result = "Pass"; // Assuming the default result is "Pass"
                    addDataToList(transposedData, "Result", result);

                    // Process headers only for the first PDF
                    if (isFirstPDF) {
                        firstPdfHeaders = transposedData.get(0);
                        writeHeadersToSheet(sheet, firstPdfHeaders);
                        isFirstPDF = false;
                    } else {
                        // Ignore headers for subsequent PDFs and use headers from the first PDF
                        transposedData.remove(0);
                    }

                    // Write the transposed data to the Excel sheet
                    writeDataToSheet(sheet, transposedData, rowIndex);
                    rowIndex += transposedData.size();
                }
            }

            // Save the combined Excel workbook
            try (FileOutputStream out = new FileOutputStream(new File(excelFilePath))) {
                workbook.write(out);
            }
        }
    }

    private static void addDataToList(List<List<String>> data, String... values) {
        for (int i = 0; i < values.length; i++) {
            if (data.size() <= i) {
                data.add(new ArrayList<>());
            }
            data.get(i).add(values[i]);
        }
    }

    private static void writeHeadersToSheet(XSSFSheet sheet, List<String> headers) {
        XSSFRow row = sheet.createRow(0);
        for (int j = 0; j < headers.size(); j++) {
            XSSFCell cell = row.createCell(j);
            cell.setCellValue(headers.get(j));
        }
    }

    private static void writeDataToSheet(XSSFSheet sheet, List<List<String>> data, int rowIndex) {
        for (int i = 0; i < data.size(); i++) {
            XSSFRow row = sheet.createRow(rowIndex + i);
            for (int j = 0; j < data.get(i).size(); j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue(data.get(i).get(j));
            }
        }
    }
}
