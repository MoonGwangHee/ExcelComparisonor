package excel.document;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;


public class Comparison {
    public static void main(String[] args) {
        try {
            String baseDir = "기본적인 루트";
            String file1Name = "기준.xls";
            String file2Name = "바꿔야함.xls";

            String file1Path = baseDir + "\\" + file1Name;
            String file2Path = baseDir + "\\" + file2Name;

            Map<String, StudentData> file1Data = readFileData(file1Path, 2, 3, 10, 8);
            Map<String, StudentData> file2Data = readFileData(file2Path, 6, 7, 10, -1);

            compareFiles(file1Data, file2Data);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Map<String, StudentData> readFileData(String filePath, int idCol, int nameCol, int scholarshipCol, int tuitionCol) throws IOException {
        Map<String, StudentData> data = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // 헤더 건너뛰기

                Cell idCell = row.getCell(idCol);
                if (idCell != null) {
                    String studentId = getCellValueAsString(idCell);
                    if (studentId.startsWith("AM")) {
                        String name = getCellValueAsString(row.getCell(nameCol)); // 이름
                        int scholarshipAmount = getIntValue(row.getCell(scholarshipCol)); // 장학금
                        int tuitionFee = (tuitionCol != -1) ? getIntValue(row.getCell(tuitionCol)) : 0; // 등록금 정보 없음

                        data.put(studentId, new StudentData(name, scholarshipAmount, tuitionFee));
                    }
                }
            }
        }
        System.out.println(filePath + " 읽은 데이터 수: " + data.size());
        return data;
    }

    private static void compareFiles(Map<String, StudentData> file1Data, Map<String, StudentData> file2Data) {
        for (Map.Entry<String, StudentData> entry : file1Data.entrySet()) {
            String studentId = entry.getKey();
            StudentData data1 = entry.getValue();
            StudentData data2 = file2Data.get(studentId);

            if (data2 != null) {
                if (data1.scholarshipAmount != data2.scholarshipAmount) {
                    System.out.println("불일치: " + studentId + " (" + data1.name + ")" +
                            " - 장학금(file1: " + data1.scholarshipAmount + ", file2: " + data2.scholarshipAmount + ")");
                }
            } else {
                System.out.println("File 2에 없는 학번: " + studentId + " (" + data1.name + ")");
            }
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf((int)cell.getNumericCellValue());
                }
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA: return cell.getCellFormula();
            default: return "";
        }
    }

    private static int getIntValue(Cell cell) {
        if (cell == null) return 0;
        switch (cell.getCellType()) {
            case NUMERIC: return (int) cell.getNumericCellValue();
            case STRING:
                try {
                    return Integer.parseInt(cell.getStringCellValue().replaceAll("[^0-9]", ""));
                } catch (NumberFormatException e) {
                    return 0;
                }
            default: return 0;
        }
    }

    private static class StudentData {
        String name;
        int scholarshipAmount;
        int tuitionFee;

        StudentData(String name, int scholarshipAmount, int tuitionFee) {
            this.name = name;
            this.scholarshipAmount = scholarshipAmount;
            this.tuitionFee = tuitionFee;
        }
    }
}
