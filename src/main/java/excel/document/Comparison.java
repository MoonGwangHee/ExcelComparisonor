package excel.document;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class Comparison {
    public static void main(String[] args) {
        try {
            String baseDir = "C:\\Users\\wldus\\OneDrive\\바탕 화면\\국가근로";
            String file1Name = "기준.xls";
            String file2Name = "바꿔야함.xls";
            String outputFileName = "output.xls";

            String file1Path = baseDir + "\\" + file1Name;
            String file2Path = baseDir + "\\" + file2Name;
            String outputPath = baseDir + "\\" + outputFileName;

            Map<String, List<Object>> file1Data = readFile1(file1Path);
            System.out.println("file1Data 크기: " + file1Data.size());
            System.out.println("file1Data 샘플: " + new ArrayList<>(file1Data.entrySet()).subList(0, Math.min(5, file1Data.size())));

            updateFile2(file2Path, outputPath, file1Data);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Map<String, List<Object>> readFile1(String filePath) throws IOException {
        Map<String, List<Object>> data = new HashMap<>();
        try (InputStream is = new FileInputStream(filePath);
             Workbook workbook = new HSSFWorkbook(is)) {
            Sheet sheet = workbook.getSheetAt(0);
            boolean inForm = false;
            int idColumnIndex = -1;
            int semesterColumnIndex = -1;
            int scholarshipColumnIndex = -1;

            for (Row row : sheet) {
                String firstCellValue = getStringCellValue(row.getCell(0));
                if (firstCellValue.equals("순번")) {
                    inForm = true;
                    // 열 인덱스 찾기
                    for (Cell cell : row) {
                        String cellValue = getStringCellValue(cell);
                        if (cellValue.equals("학번")) idColumnIndex = cell.getColumnIndex();
                        if (cellValue.equals("학기")) semesterColumnIndex = cell.getColumnIndex();
                        if (cellValue.equals("장학금")) scholarshipColumnIndex = cell.getColumnIndex();
                    }
                    continue;
                }

                if (inForm && !firstCellValue.startsWith("소계") && !firstCellValue.isEmpty()) {
                    String studentId = getStringCellValue(row.getCell(idColumnIndex));
                    String semester = getStringCellValue(row.getCell(semesterColumnIndex));
                    String scholarship = getStringCellValue(row.getCell(scholarshipColumnIndex));

                    if (studentId.startsWith("AM")) {
                        List<Object> rowData = new ArrayList<>();
                        rowData.add(semester);
                        rowData.add(scholarship);
                        data.put(studentId.trim(), rowData);
                    }
                }

                if (firstCellValue.startsWith("소계")) {
                    inForm = false;
                }
            }
        }
        return data;
    }

    private static void updateFile2(String inputPath, String outputPath, Map<String, List<Object>> file1Data) throws IOException {
        try (InputStream is = new FileInputStream(inputPath);
             Workbook workbook = new HSSFWorkbook(is);
             OutputStream os = new FileOutputStream(outputPath)) {
            Sheet sheet = workbook.getSheetAt(0);

            boolean inForm = false;
            int idColumnIndex = -1;
            int yearColumnIndex = -1;
            int scholarshipColumnIndex = -1;

            for (Row row : sheet) {
                String firstCellValue = getStringCellValue(row.getCell(0));
                if (firstCellValue.equals("No")) {
                    inForm = true;
                    // 열 인덱스 찾기
                    for (Cell cell : row) {
                        String cellValue = getStringCellValue(cell);
                        if (cellValue.equals("학번")) idColumnIndex = cell.getColumnIndex();
                        if (cellValue.equals("학년")) {
                            cell.setCellValue("학기");
                            yearColumnIndex = cell.getColumnIndex();
                            System.out.println("'학년' 열을 '학기'로 변경했습니다. 인덱스: " + yearColumnIndex);
                        }
                        if (cellValue.equals("장학금액")) scholarshipColumnIndex = cell.getColumnIndex();
                    }
                    continue;
                }

                if (inForm && !firstCellValue.isEmpty() && !firstCellValue.contains("장학종별 소계")) {
                    Cell idCell = row.getCell(idColumnIndex);
                    if (idCell != null) {
                        String studentId = getStringCellValue(idCell).trim();
                        System.out.println("처리 중인 학번: " + studentId);

                        if (file1Data.containsKey(studentId)) {
                            List<Object> file1RowData = file1Data.get(studentId);
                            System.out.println("file1에서 찾은 데이터: " + file1RowData);

                            // 학기 정보 업데이트
                            Cell semesterCell = row.getCell(yearColumnIndex);
                            String file1Semester = (String) file1RowData.get(0);
                            String currentSemester = getStringCellValue(semesterCell);

                            System.out.println("file1 학기: " + file1Semester + ", 현재 file2 학기: " + currentSemester);

                            if (!currentSemester.equals(file1Semester)) {
                                semesterCell.setCellValue(file1Semester);
                                System.out.println("학기 수정됨: " + studentId + " - 원래 값: " + currentSemester + ", 새 값: " + file1Semester);
                            } else {
                                System.out.println("학기가 이미 일치합니다. 수정 불필요.");
                            }

                            // 장학금 정보 업데이트
                            Cell scholarshipCell = row.getCell(scholarshipColumnIndex);
                            String file1Scholarship = (String) file1RowData.get(1);
                            String currentScholarship = getStringCellValue(scholarshipCell);
                            if (!currentScholarship.equals(file1Scholarship)) {
                                scholarshipCell.setCellValue(file1Scholarship);
                                System.out.println("장학금 수정됨: " + studentId + " - 원래 금액: " + currentScholarship + ", 새 금액: " + file1Scholarship);
                            }
                        } else {
                            System.out.println("학번 " + studentId + "는 file1에 없습니다.");
                        }
                    }
                }

                if (firstCellValue.contains("장학종별 소계")) {
                    inForm = false;
                }
            }

            workbook.write(os);
            System.out.println("파일 업데이트가 완료되었습니다.");
        }
    }

    private static String getStringCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf((int) cell.getNumericCellValue());
                }
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA: return cell.getCellFormula();
            default: return "";
        }
    }
}
