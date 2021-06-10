package qwr;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.awt.Color;

import static org.apache.poi.hsmf.datatypes.Types.NULL;

public class Stylq {
    static int cstl=8; //количество стилей

    static XSSFCellStyle creatStyle(int i){
        XSSFFont font = Main.wbook.createFont();
//        font.setFontName("Courier New");
        font.setItalic(false);//наклонный
        switch (i){
            case 1://заголовок документа
                font.setFontHeightInPoints((short)14);
                font.setBold(true);//жирный
                break;
            case 2://шапка таблицы
                font.setFontHeightInPoints((short)8);
                font.setBold(false);//жирный
                break;
            case 3://группировка
                font.setFontHeightInPoints((short)10);
                font.setBold(false);//жирный
                break;
            case 7://скрытый текст
                font.setColor(new XSSFColor(new Color(210, 233, 243)));
            default:
                font.setFontHeightInPoints((short)10);
                font.setBold(false);//жирный
        }//switch (i)
        XSSFCellStyle style= Main.wbook.createCellStyle();
        style.setWrapText(true);//перенос текста в ячейке
        style.setFont(font);
        switch (i){//работаю без бордюра
            case 1: break;
            case 3: break;
            default:
                style.setBorderRight (BorderStyle.HAIR );
                style.setBorderBottom(BorderStyle.HAIR );
                style.setBorderLeft  (BorderStyle.HAIR );
                style.setBorderTop   (BorderStyle.HAIR );
        }//switch
        style.setVerticalAlignment(VerticalAlignment.CENTER);//сработало по вертикал
        switch (i) {
            case 2://шапка таблицы
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
            case 1://заголовок документа

            case 5://номер строки Integer
                style.setAlignment(HorizontalAlignment.CENTER);//по горизонтали
                break;
            case 3://группировка
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);//
            case 4://text
                style.setAlignment(HorizontalAlignment.LEFT);//горизонтали
                break;
            default:// 6) Double
                style.setAlignment(HorizontalAlignment.RIGHT);//горизонтали
                CreationHelper createHelper = Main.wbook.getCreationHelper();
            style.setDataFormat(createHelper.createDataFormat().getFormat("#.##"));
        }//switch
        return style;
    }//creatStyle
    //------------------------------------------------------------

    public static int clearTitle(CellStyle cs,int fr,int lr,int fc,int lc,String s){
        for (int i=fr; i<= lr; i++) {
            Main.sheet.createRow((short)i);
            for (int j = fc; j <= lc; j++)
                Main.sheet.getRow(i).createCell(j).setCellStyle(cs);
        }
        Main.sheet.getRow(fr).getCell(fc).setCellValue(s);
        Main.sheet.addMergedRegion(new CellRangeAddress(fr,lr,fc, lc) );
        return lr+1;
    }//clearTitle
    //-------------------------------------------------------------

    public static void addTitle(CellStyle cs,int fr,int lr,int fc,int lc, String s){
        for (int i=fr; i<= lr; i++) {
            for (int j = fc; j <= lc; j++)
                Main.sheet.getRow(i).createCell(j).setCellStyle(cs);
        }
        Main.sheet.getRow(fr).getCell(fc).setCellValue(s);
        Main.sheet.addMergedRegion(new CellRangeAddress(fr,lr,fc, lc) );
    }//addTitle
    //--------------------------------------------------------------
    /**
     *
     * @param cs задает стиль, которым пишется ячейка
     * @param fr указывает номер строки куда писать
     * @param fc указывает номер столбца куда писать
     * @param s  запись самой формулы
     */
    public static void addFormlN(CellStyle cs, int fr, int fc,  String s){
        Main.sheet.getRow(fr).createCell(fc).setCellFormula(s);
        Main.sheet.getRow(fr).getCell(fc).setCellStyle(cs);
        String formula = Main.sheet.getRow(fr).getCell(fc).getCellFormula();
FormulaEvaluator evaluator = Main.wbook.getCreationHelper().createFormulaEvaluator();
        CellValue cellValue = evaluator.evaluate(Main.sheet.getRow(fr).getCell(fc));
        double value = cellValue.getNumberValue();
        Main.sheet.getRow(fr).getCell(fc).setCellValue(value);
        //        String value = cellValue.getStringValue();
        //        boolean value = cellValue.getBooleanValue();
    }//addForml
}//class