package syuri;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sample {

    public static final String CONST_STR_FILE_ROOT = "C:\\test\\";
    public static final String CONST_STR_ORIGINAL_FILENAME = "sample1.xlsx";
    public static final String CONST_STR_ORIGINAL_SHEET_NAME = "SheetName";// "受付台帳";
    public static final String CONST_STR_NEW_FILENAME = "sample2_1.xlsx";
    public static final String CONST_STR_NEW_SHEET_NAME = "nWbSheet";
    public static final String CONST_STR_EMPTY = "";
    public static final String CONST_STR_SPACE = " ";
    public static final String CONST_STR_KAIGYO = "\r\n";

    public static void main(String[] args) {
        FileInputStream in = null;
        Workbook oWb = null;
        Workbook nWb = null;
        try {
            String path = "";
            in = new FileInputStream(CONST_STR_FILE_ROOT + CONST_STR_ORIGINAL_FILENAME);
            oWb = WorkbookFactory.create(in);
            nWb = new XSSFWorkbook();// Excel2007~
        } catch (IOException e) {
            System.out.println(e.toString());
        } catch (InvalidFormatException e) {
            System.out.println(e.toString());
        } finally {
            try {
                in.close();
            } catch (IOException e) {
                System.out.println(e.toString());
            }
        }

        // 初期設定
        Sheet oSheet = oWb.getSheet(CONST_STR_ORIGINAL_SHEET_NAME);
        Sheet nSheet = nWb.createSheet(CONST_STR_NEW_SHEET_NAME);

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();
        CreationHelper createHelper = nWb.getCreationHelper();
        CellStyle cellStyle = nWb.createCellStyle();
        short style = createHelper.createDataFormat().getFormat("yyyy/mm/dd");
        cellStyle.setDataFormat(style);

        for (Row oRow : oSheet) {

            if (oRow != null) {
                int row = oRow.getRowNum();
                if (row < 2) {
                    continue;
                }
                Row nRow = nSheet.createRow(row);
                // String receptionNo = oRow.getCell(0).getStringCellValue();

                for (Cell oCell : oRow) {
                    Cell nCell = nRow.createCell(oCell.getColumnIndex());
                    int cellTypeNum = oCell.getCellType();
                    System.out.print(cellTypeNum);
                    switch (cellTypeNum) {
                    case Cell.CELL_TYPE_NUMERIC:
                        if (DateUtil.isCellDateFormatted(oCell)) {
                            nCell.setCellStyle(cellStyle);
                            nCell.setCellValue(DateUtil.getJavaDate(oCell.getNumericCellValue()));// DateUtil.getJavaDate(oCell.getNumericCellValue()).toString());
                        } else {
                            nCell.setCellValue(oCell.getNumericCellValue());
                        }
                        break;
                    case Cell.CELL_TYPE_STRING:
                        nCell.setCellValue(oCell.getStringCellValue());
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        try {
                            nCell.setCellFormula(oCell.getCellFormula());
                            // nCell.setCellValue(oCell.getStringCellValue());
                        } catch (IllegalStateException e) {
                            System.out.println(e.toString());
                        }

                        // nCell.setCellValue(oCell.getStringCellValue());
                        // nCell.setCellFormula(oCell.getCellFormula());
                        break;
                    case Cell.CELL_TYPE_BLANK:
                        nCell.setCellType(cellTypeNum);
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        nCell.setCellType(cellTypeNum);
                        nCell.setCellValue(oCell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_ERROR:
                        nCell.setCellType(cellTypeNum);
                        break;
                    default:
                        break;
                    }
                    // nCell.setCellValue(oCell.toString());
                    // nCell = oCell;
                    System.out.print("[" + oCell.getRowIndex() + ":" + oCell.getColumnIndex() + "] = " + oCell
                            + CONST_STR_KAIGYO);
                }
            }
        }

        FileOutputStream out = null;
        try {
            out = new FileOutputStream(CONST_STR_FILE_ROOT + CONST_STR_NEW_FILENAME);
            nWb.write(out);
        } catch (IOException e) {
            System.out.println(e.toString());
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                System.out.println(e.toString());
            }
        }
    }
}
