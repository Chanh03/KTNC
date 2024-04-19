import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelWriter {
    public static void writeToExcel(String excelDir, String sheetName, Map<String, Object[]> data) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(sheetName);

        try {
            if (data != null) {
                // Tạo header cho bảng kết quả
                Row headerRow = sheet.createRow(0);
                String[] headers = {"ID", "Dữ liệu test", "Các bước test", "Kết quả thực tế", "Kết quả mong muốn", "Trạng thái ", "Thời gian"};
                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers[i]);
                }

                // Ghi dữ liệu từ hàng thứ 2 trở đi
                int rownum = 1; // Bắt đầu từ dòng thứ 2
                for (Object[] objArr : data.values()) {
                    Row row = sheet.createRow(rownum++);
                    int cellnum = 0;
                    for (Object obj : objArr) {
                        Cell cell = row.createCell(cellnum++);
                        if (obj instanceof String) {
                            cell.setCellValue((String) obj);
                        } else if (obj instanceof LocalDateTime) {
                            cell.setCellValue(obj.toString());
                        } else {
                            cell.setCellValue(String.valueOf(obj));
                        }
                    }
                }

                // Lưu workbook vào file Excel
                FileOutputStream out = new FileOutputStream(new File(excelDir));
                workbook.write(out);
                out.close();
                System.out.println("Kết quả kiểm thử đã được ghi vào: " + excelDir);
            } else {
                System.out.println("Không có kết quả kiểm thử để ghi.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
