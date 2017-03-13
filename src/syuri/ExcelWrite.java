package syuri;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
    //適当なディレクトリに書き換えてください
    static final String INPUT_DIR = "C:\\test\\output\\";

    public static void main(String[] args) {

        try {

            //xlsの場合はこちらを有効化
            //Workbook wb = new HSSFWorkbook();
            //FileOutputStream fileOut = new FileOutputStream("workbook.xls");

            //xlsxの場合はこちらを有効化
            Workbook wb = new XSSFWorkbook();
            FileOutputStream fileOut = new FileOutputStream(INPUT_DIR + "sample2.xlsx");

            String safeName = WorkbookUtil.createSafeSheetName("['aaa's test*?]");
            Sheet sheet1 = wb.createSheet(safeName);

            CreationHelper createHelper = wb.getCreationHelper();

            //Rows(行にあたる)を作る。Rowsは0始まり。
            Row row = sheet1.createRow((short)0);
            //cell(列にあたる)を作って、そこに値を入れる。
            Cell cell = row.createCell(0);
            cell.setCellValue(1);

            row.createCell(1).setCellValue(1.2);
            row.createCell(2).setCellValue(
                 createHelper.createRichTextString("sample string"));
            row.createCell(3).setCellValue(true);

            wb.write(fileOut);
            fileOut.close();

        }catch (Exception e) {
            e.printStackTrace();
        } finally {
        }
    }

}
