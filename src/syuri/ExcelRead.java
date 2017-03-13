package syuri;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

    //適当なディレクトリに書き換えてください
    static final String INPUT_DIR = "C:\\test\\";

    public static void main(String[] args) {

        
        try {
            String xlsxFileAddress = INPUT_DIR + "Sample1.xlsx";
            //共通インターフェースを扱える、WorkbookFactoryで読み込む
            Workbook wb = WorkbookFactory.create(new FileInputStream(xlsxFileAddress));
            //全セルを表示する
            for (Sheet sheet : wb ) {
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        System.out.print(getCellValue(cell));
                        System.out.print(" , ");
                    }
                    System.out.println();
                }
            }
            wb.close();
        }catch (Exception e) {
            e.printStackTrace();
        } finally {
        }
    }
    
    private static Object getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                return cell.getRichStringCellValue().getString();
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_FORMULA:
                return cell.getCellFormula();
            default:
                return null;
        }
    }
    
    public static void ExcelWrite (String[] args) {
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

