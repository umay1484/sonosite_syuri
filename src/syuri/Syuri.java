package syuri;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

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

public class Syuri {

    public static final String CONST_STR_FILE_ROOT = "file\\";// "C:\\test\\";
    public static final String CONST_STR_ORIGINAL_FILENAME = "2017受付台帳Ver2.6.xls";
    public static final String CONST_STR_ORIGINAL_SHEET_NAME = "受付台帳";// "受付台帳";
    public static final String CONST_STR_NEW_FILENAME = "sample2_1.xlsx";
    public static final String CONST_STR_NEW_SHEET_NAME = "nWbSheet";
    public static final String CONST_STR_EMPTY = "";
    public static final String CONST_STR_SPACE = " ";
    public static final String CONST_STR_KAIGYO = "\r\n";

    public static Sheet chkSheet(String fileRoot, String fileName, String sheetName) {
        FileInputStream tmpIn = null;
        Workbook tmpWb = null;
        Sheet tmpSheet = null;
        try {
            tmpIn = new FileInputStream(fileRoot + fileName);
            // OriginalExcel
            tmpWb = WorkbookFactory.create(tmpIn);
            tmpSheet = tmpWb.getSheet(sheetName);

        } catch (IOException e) {
            System.out.println(e.toString());
        } catch (InvalidFormatException e) {
            System.out.println(e.toString());
        } finally {
            try {
                tmpIn.close();
                return tmpSheet;
            } catch (IOException e) {
                System.out.println(e.toString());
            }
        }
        return null;
    }

    public static Cell setCell(int tmpCellTypeNum, Cell oTmpCell, Cell nTmpCell, CellStyle tmpCellStyle) {

        switch (tmpCellTypeNum) {
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(oTmpCell)) {
                    nTmpCell.setCellStyle(tmpCellStyle);
                    nTmpCell.setCellValue(DateUtil.getJavaDate(oTmpCell.getNumericCellValue()));// DateUtil.getJavaDate(oCell.getNumericCellValue()).toString());
                } else {
                    nTmpCell.setCellValue(oTmpCell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_STRING:
                nTmpCell.setCellValue(oTmpCell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                try {
                    nTmpCell.setCellValue(oTmpCell.getStringCellValue());
                    // nCell.setCellValue(oCell.getStringCellValue());
                } catch (IllegalStateException e) {
                    // nCell.setCellFormula(oCell.getCellFormula());
                    nTmpCell.setCellValue(e.toString());
                    System.out.println(e.toString());
                }
                // nCell.setCellValue(oCell.getStringCellValue());
                // nCell.setCellFormula(oCell.getCellFormula());
                break;
            case Cell.CELL_TYPE_BLANK:
                nTmpCell.setCellType(tmpCellTypeNum);
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                nTmpCell.setCellType(tmpCellTypeNum);
                nTmpCell.setCellValue(oTmpCell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_ERROR:
                nTmpCell.setCellType(tmpCellTypeNum);
                break;
            default:
                break;
        }
        return nTmpCell;
    }

    public static void main(String[] args) {
        String oFileNameList[] = { "2012受付台帳Ver2.2.xls", "2013受付台帳Ver2.3.xls", "2014受付台帳Ver2.3.xls", "2015受付台帳Ver2.4.xls", "2016受付台帳Ver2.5.xls", "2017受付台帳Ver2.6.xls" };
        int oColNumList[] = { 0, 1, 2, 4, 7, 12, 14 };
        String oSheetNameList[] = { "受付台帳2012", "受付台帳", "受付台帳", "受付台帳", "受付台帳", "受付台帳" };
        int oRowNumStartList[] = { 3, 4, 4, 4, 4, 4 };
        // FileInputStream in = null;
        // Workbook oWb = null;
        Sheet oSheet = null;
        Workbook nWb = null;
        Sheet nSheet = null;
        CellStyle cellStyle = null;
        Cell tmpCell = null;

        short style = 0;
        int rowNum = 0;
        int colNum = 0;
        int oRowNumStart = 2;
        int oColNumDecision = 7;// 元ファイルの行判定カラム：修理完了日
        // 集計元シートの集計対象カラム
        int oReceptionNumberColumn = 0;
        int oSheetNameCloumn = 1;
        int oSrNumberCloumn = 2;
        int oRequestProcessCloumn = 4;
        int oRepairCompleteDateCloumn = 7;
        int oModelCloumn = 12;
        int oSerialNumberCloumn = 13;
        int oRefNumberCloumn = 14;
        int oPartsNumberCloumn = 26;
        int oPartsNumberrCount = 10;
        List<Integer> oCloumnList = new ArrayList<Integer>();
        oCloumnList.add(oReceptionNumberColumn);
        oCloumnList.add(oSheetNameCloumn);
        oCloumnList.add(oSrNumberCloumn);
        oCloumnList.add(oRequestProcessCloumn);
        oCloumnList.add(oRepairCompleteDateCloumn);
        oCloumnList.add(oModelCloumn);
        oCloumnList.add(oSerialNumberCloumn);
        oCloumnList.add(oRefNumberCloumn);
        for (int i = 0; i < oPartsNumberrCount; i++) {
            oCloumnList.add(oPartsNumberCloumn + i);
        }
        // 出力シート
        String nReceptionNumber = null;
        String nSheetName = null;
        String nSrNumber = null;
        String nRequestProcess = null;
        double nRepairCompleteDate = 0;
        String nModel = null;
        String nSerialNumber = null;
        String nPartsNumberList[] = null;
        int nPartsCount = 0;
        int cellTypeNumber = 0;

        nWb = new XSSFWorkbook();// Excel2007~
        nSheet = nWb.createSheet(CONST_STR_NEW_SHEET_NAME);
        // 日付書式設定
        CreationHelper createHelper = nWb.getCreationHelper();
        cellStyle = nWb.createCellStyle();
        style = createHelper.createDataFormat().getFormat("yyyy/mm/dd");
        cellStyle.setDataFormat(style);

        for (int i = 0; i < oFileNameList.length; i++) {
            oRowNumStart = oRowNumStartList[i];
            oSheet = chkSheet(CONST_STR_FILE_ROOT, oFileNameList[i], oSheetNameList[i]);

            for (Row oRow : oSheet) {
                // スキップ
                if (oRow == null) {
                    continue;
                }
                if (oRow.getRowNum() < oRowNumStart) {
                    continue;
                }
                if (oRow.getCell(oSheetNameCloumn) == null) {
                    continue;
                } else if (oRow.getCell(oSheetNameCloumn).getCellType() == Cell.CELL_TYPE_BLANK) {
                    continue;
                } else if (oRow.getCell(oSheetNameCloumn).getStringCellValue().substring(0, 3).indexOf("代替機") >= 0
                        || oRow.getCell(oSheetNameCloumn).getStringCellValue().substring(0, 3).indexOf("デモ機") >= 0) {
                    System.out.println("代替機||デモ機");
                    continue;
                }
                if (oRow.getCell(oRepairCompleteDateCloumn) == null) {
                    continue;
                } else if (oRow.getCell(oRepairCompleteDateCloumn).getCellType() == Cell.CELL_TYPE_FORMULA) {
                    continue;
                }

                // 集計対象
                Row nRow = nSheet.createRow(rowNum);
                colNum = 0;
                System.out.print(nRow.getRowNum() + ": ");
                int j = 0;
                outputCell: while (j < oCloumnList.size()) {
                    if (oRow.getCell(oCloumnList.get(j)) == null) {
                        break;
                    } else {
                        cellTypeNumber = oRow.getCell(oCloumnList.get(j)).getCellType();
                        Cell nCell = nRow.createCell(j);
                        switch (cellTypeNumber) {
                            case Cell.CELL_TYPE_NUMERIC:
                                try {
                                    if (DateUtil.isCellDateFormatted(oRow.getCell(oCloumnList.get(j)))) {
                                        nCell.setCellStyle(cellStyle);
                                        nCell.setCellValue(DateUtil.getJavaDate(oRow.getCell(oCloumnList.get(j)).getNumericCellValue()));
                                    } else {
                                        nCell.setCellValue(oRow.getCell(oCloumnList.get(j)).getNumericCellValue());
                                    }
                                } catch (Exception e) {
                                    System.out.println(e.toString());
                                    nSheet.removeRowBreak(rowNum);
                                    rowNum--;
                                    break outputCell;
                                }
                                break;
                            case Cell.CELL_TYPE_STRING:
                                try {
                                    nCell.setCellValue(oRow.getCell(oCloumnList.get(j)).getStringCellValue());
                                } catch (Exception e) {
                                    System.out.println(e.toString());
                                    nSheet.removeRowBreak(rowNum);
                                    rowNum--;
                                    break outputCell;
                                }
                                break;
                            case Cell.CELL_TYPE_FORMULA:
                                try {
                                    nCell.setCellValue(oRow.getCell(oCloumnList.get(j)).getStringCellValue());
                                } catch (IllegalStateException e) {
                                    // nCell.setCellFormula(oCell.getCellFormula());
                                    System.out.println(e.toString());
                                    nSheet.removeRowBreak(rowNum);
                                    rowNum--;
                                    break outputCell;
                                } catch (Exception e) {
                                    System.out.println(e.toString());
                                    nSheet.removeRowBreak(rowNum);
                                    rowNum--;
                                    break outputCell;
                                }
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                break;
                            case Cell.CELL_TYPE_BOOLEAN:
                                break;
                            case Cell.CELL_TYPE_ERROR:
                                break;
                            default:
                                break;
                        }
                        System.out.print(nCell + ", ");
                    }
                    j++;
                }
                System.out.println();
                rowNum++;
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
}
