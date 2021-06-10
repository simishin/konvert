package qwr;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import qwr.utl.BgFile;

import static qwr.utl.BgFile.prnq;

public class Main {
    static XSSFWorkbook wbook;
    static XSSFSheet sheet;
    public static String filenameoutw ="werq.xlsx";//имя файла результата
    public static String filenameoutv ="verq.xlsx";//имя файла результата
    public static String filenamein;//имя файла результата
    public static ArrayList<Elmnt> elst=new ArrayList<Elmnt>(220);//список входнвых данных
    //--------------------------------------------------------------
    public static boolean prnq(String s){System.out.println(s); return true;}
    public static boolean prnt(String s){System.out.print(s); return true;}
    //---------------------------------------------------------------
    public static void main(String[] args) {
        if(!BgFile.comstring(args))System.exit(78); //новая инициалзация

        prnq("Читаю файл заготовки темплана "+ BgFile.flin);
        Elmnt q=new Elmnt();// элемент колекции для подготовки ввода
//        prnq("читаю файл "+filenamein);
        short rowi=0;//счетчик строк
        int celi=0;//индекс ячейки
        try(FileInputStream file = new FileInputStream(new File(BgFile.flin)))
        {
            int nerr=0;//подсчет количества ошибок
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            label:
            while (rowIterator.hasNext())//цикл по строкам
            {
                Row row = rowIterator.next();
                assert prnq("");
                assert prnt((row.getRowNum()+1)+") ");
                if (rowi++ > 10) {
                    q.clear();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    //пропускаю первые строки
                    while (cellIterator.hasNext())//цикл по ячейкам строки
                    {
                        Cell cell = cellIterator.next();
                        celi= cell.getColumnIndex();//индекс ячейки
//                        prnt(" "+celi);
                        CellType cellType = cell.getCellType();
                        switch (cellType) {
                            case ERROR: prnq(" "+celi+"ERROR "); nerr++; break;
//                               case _NONE: assert prnt(" _NONE "); break;
//                               case BLANK: assert prnt("ъ "); break;
                            case FORMULA:
//                                    assert prnt("f");
                                switch (cell.getCachedFormulaResultType()){
                                    case NUMERIC:
                                        assert prnt("N ");
                                        q.setdbl(celi, cell.getNumericCellValue());
                                        break;
                                    case STRING: assert prnt("S ");break;
                                    case ERROR:   assert prnt("E ");break;
                                    case BLANK:   assert prnt("B ");break;
                                    case _NONE:   assert prnt("O ");break;
                                }
                                break;//завершение анализа формулы
                            case NUMERIC://число
                                q.setdbl(celi, cell.getNumericCellValue());
                                assert prnt("n ");
                                break;
                            case STRING://строка
                                q.setstr(celi, cell.getStringCellValue());
                                assert prnt("s ");
                                break;
                            default: assert prnt("@ ");
                        }//switch
                    }//while по строке

                    if (q.isAdd() ) {if (! (elst.add ( new Elmnt(q)))) prnq("ощибка сохранения в коллекции");}
                }
            }//while по листу
            file.close();
            prnq("");
            prnq("Входной файл прочитан в количестве строк "+rowi);
            prnq("Количество ячееек с ошибкой "+nerr);
        }
        catch (Exception e){ e.printStackTrace(); }
        finally {        }
        if (elst.size()<2) {prnq("Строки не распознаны! ");System.exit(78);}
        else prnq("Количество распознаных строк "+elst.size());
        //-------------------------------------------------------------------------------
        assert prnq(" *2* формирования перечня исполнителей* ");
        HashSet<String> lpipl= new HashSet<String>();
        for (Elmnt itn : elst){ lpipl.add(itn.getPipl()); }
//        for (String s: lpipl) prnq(" "+s);  prnq("****");

        assert prnq(" *3* начало формирования ежемесячного отчета* ");
        ArrayList<Elcol> wtu = new ArrayList<>(36);

//        wtu.add(new Elcol(1200,"№ п.п.",2,14,4));

        wtu.add(new Elcol(1100,"№№",1,14,4));
        wtu.add(new Elcol(7000,"№ сметы",1,5,3));
        wtu.add(new Elcol(12500,"наименование работы",1,2,3));
        wtu.add(new Elcol(2000,"Сметная стоимость в базовых ценах 2000г., руб.",1,11,5));
        wtu.add(new Elcol(2000,"остаток на начало периода",1,12,5));
        wtu.add(new Elcol(2000,"план на период",1,13,5));
        wtu.add(new Elcol(2000,"факт за период",1,15,5));
        wtu.add(new Elcol(2000,"Исполнитель",1,6,3));
        wtu.add(new Elcol(5200,"примечание",1,16,3));
        //*******************
        int colsz=wtu.size()-1;
        prnq("Количество колонок "+(colsz+1));

        wbook = new XSSFWorkbook();//создаю книгу
        //создаю массив стилей
        XSSFCellStyle[] slt=new XSSFCellStyle[Stylq.cstl];
        for (int i=0; i<Stylq.cstl; i++) slt[i]=Stylq.creatStyle((i+1));

        sheet = wbook.createSheet("Reports");//создаю лист
        for (Elcol i: wtu) sheet.setColumnWidth(i.getNcol(),i.getWid());

        sheet.addMergedRegion(new CellRangeAddress(0,0,0, wtu.size()-1) );

        rowi=0;//счетчик строк
        XSSFRow rowx = sheet.createRow(rowi++);//создаем строку
        rowx.setHeight((short) 1295);//высота строки
        XSSFCell cell = rowx.createCell(0);

        cell.setCellStyle( slt[0]);
        cell.setCellType(CellType.STRING);
        cell.setCellValue("ООО \"УС БАЭС\" \n  за период с 16.01.2020г.по16.02.2020г. \n \"Строительство модуля фабрикации и пускового комплекса рефабрикации плотного смешанного уранплутониевого топлива для реакторов на быстрых нейтронах\" ");
        rowx = sheet.createRow(rowi++);//создаем строку
        rowx.setHeight((short) 885);//высота строки
        String formulaq;
//        XSSFCellStyle styleB=Stylq.creatStyle(2);
        for (Elcol i: wtu) {
            rowx.createCell(i.getNcol()).setCellValue(i.getName());
            rowx.getCell(i.getNcol()).setCellStyle(slt[1]);
        }//for wtu
        //-------------------------------------------------------------------
        for (Elmnt itn : elst){
            rowx = sheet.createRow(rowi);//создаем строку

            if (itn.getIdr()>=0) {
                for (Elcol i: wtu) {
                    if(itn.getVal(i.getNdt()) instanceof String){//является ли объект строкой
                        rowx.createCell(i.getNcol()).setCellValue((String)itn.getVal(i.getNdt()));
                        rowx.getCell(i.getNcol()).setCellStyle(slt[i.getStl()]);
                    }
                    else if(itn.getVal(i.getNdt()) instanceof Integer){//является ли объект числом
                        rowx.createCell(i.getNcol()).setCellValue((Integer)itn.getVal(i.getNdt()));
                        rowx.getCell(i.getNcol()).setCellStyle(slt[i.getStl()]);
                    }
                    else if(itn.getVal(i.getNdt()) instanceof Double){//является ли объект числом
                        rowx.createCell(i.getNcol()).setCellValue((Double)itn.getVal(i.getNdt()));
                        rowx.getCell(i.getNcol()).setCellStyle(slt[i.getStl()]);
                    }
                }//for wtu
                //добавляю формулу за пределы строки
                formulaq="SUM(L"+(rowi+1)+":U"+(rowi+1)+")";
                Stylq.addFormlN(slt[5],rowi,10,formulaq);
            } else {
                cell = rowx.createCell(0);
                cell.setCellStyle(slt[2]);//установка стиля на ячейку
                cell.setCellType(CellType.STRING);
                cell.setCellValue(itn.getName());
                sheet.addMergedRegion(new CellRangeAddress(rowi,rowi,0, 8) );
            }//if
            rowi++;
        }//for elst (row)
        prnq(cell.getRawValue()+" "+cell.getReference()+" "+(rowi-1));
        //дописываю формулу и сразу далаю по ней расчет
       formulaq="SUM(K3:K"+(rowi+30)+")/1000";
        rowx = sheet.createRow(rowi);//создаем строку
        Stylq.addFormlN(slt[5],1,10,formulaq);

        FormulaEvaluator evaluator;
        CellValue cellValue;
        double value;
        BgFile.savefl(wbook,filenameoutw);//сохранение результата работы
//--------------------------------------------------------------------------
        assert prnq(" *4* формирование документа ежедневного отчета* ");
        int sk=5;
        ArrayList<Elcol> lecol = new ArrayList<>(36);
        Elcol.counter=0;
        lecol.add(new Elcol(1200,"№ п.п.",2,14,4));
        lecol.add(new Elcol(2300,"Исполнитель",2,6,3));
        lecol.add(new Elcol(8000,"Наименование работ",2,2,3));
        lecol.add(new Elcol(2000,"№ сметы",3,4,3));
        lecol.add(new Elcol(2000,"Инв № чертежа",3,3,3));
        lecol.add(new Elcol(2000,"Ед. изм.",3,7,4));
        lecol.add(new Elcol(2000,"Всего по проекту",3,8,sk));
        lecol.add(new Elcol(2000,"Отстаок на 15.10.н",3,9,sk));
        lecol.add(new Elcol(1500,"План",4,10,sk));
        lecol.add(new Elcol(1500,"Факт",4,0,sk,"R9+V9+Z9+AD9"));
        lecol.add(new Elcol(1500,"Отклонение",4,0,sk,"Q9+U9-R9-V9"));
        lecol.add(new Elcol(2000,"План",4,17,sk));//пллан стоимости на период в текущих
        lecol.add(new Elcol(2000,"Факт",4,0,sk,"T9+X9+AB9+AF9"));
        lecol.add(new Elcol(2000,"Отклонение",4,0,sk,"P9-O9"));
        lecol.add(new Elcol(2000,"План",4,0,sk,"S9"));
        lecol.add(new Elcol(2000,"Факт",4,0,sk,"T9"));
        lecol.add(new Elcol(1500,"План",4,15,sk));
        lecol.add(new Elcol(1500,"Факт",4,15,sk));
        lecol.add(new Elcol(2000,"План",4,0,sk,"$L9/$I9*Q9"));
        lecol.add(new Elcol(2000,"Факт",4,0,sk,"$L9/$I9*R9"));
        lecol.add(new Elcol(1500,"План",4,15,sk));
        lecol.add(new Elcol(1500,"Факт",4,15,sk));
        lecol.add(new Elcol(2000,"План",4,0,sk,"$L9/$I9*U9"));
        lecol.add(new Elcol(2000,"Факт",4,0,sk,"$L9/$I9*V9"));
        lecol.add(new Elcol(1500,"План",4,15,sk));
        lecol.add(new Elcol(1500,"Факт",4,15,sk));
        lecol.add(new Elcol(2000,"План",4,0,sk,"$L9/$I9*Y9"));
        lecol.add(new Elcol(2000,"Факт",4,0,sk,"$L9/$I9*Z9"));
        lecol.add(new Elcol(1500,"План",4,15,sk));
        lecol.add(new Elcol(1500,"Факт",4,15,sk));
        lecol.add(new Elcol(2000,"План",4,0,sk,"$L9/$I9*AC9"));
        lecol.add(new Elcol(2000,"Факт",4,0,sk,"$L9/$I9*AD9"));

        lecol.add(new Elcol(3000,"Причины отклонений от плана",2,16,4));
        wbook = new XSSFWorkbook();//создаю книгу
        //создаю массив стилей
        slt=new XSSFCellStyle[Stylq.cstl];//повторное назначение
        for (int i=0; i<Stylq.cstl; i++) slt[i]=Stylq.creatStyle((i+1));

        sheet = wbook.createSheet("Reports");//создаю лист
        for (Elcol i: lecol) sheet.setColumnWidth(i.getNcol(),i.getWid());
        sheet.addMergedRegion(new CellRangeAddress(0,0,3, 11) );
        rowi=0;//счетчик строк
        rowx = sheet.createRow(rowi++);//создаем строку
        rowx.setHeight((short) 995);//высота строки
        sheet.getRow(0).createCell(3).setCellValue("Еженедельный отчет по ТП выполения СМР МФР за период июль-август 2020 ");
        sheet.getRow(0).getCell(3).setCellStyle(slt[0]);
    //    rowx.createCell(3).setCellValue("Еженедельный отчет по ТП выполения СМР МФР за период июль-август 2020 ");
//        rowx.getCell(3).setCellStyle(slt[0]);
        rowx = sheet.createRow(rowi++);//создаем строку
        sheet.addMergedRegion(new CellRangeAddress(1,1,3, 11) );
        rowx.setHeight((short) 395);//высота строки
        sheet.getRow(1).createCell(3).setCellValue("По состоянию на 16.10.2020");
        sheet.getRow(1).getCell(3).setCellStyle(slt[0]);
        rowx = sheet.createRow(rowi++);//создаем строку
        sheet.addMergedRegion(new CellRangeAddress(2,2,3, 4) );
        sheet.addMergedRegion(new CellRangeAddress(2,2,5, 10) );
        sheet.addMergedRegion(new CellRangeAddress(2,3,11, 13) );
        sheet.addMergedRegion(new CellRangeAddress(2,3,14, 15) );
        sheet.addMergedRegion(new CellRangeAddress(2,2,16, 19) );
        sheet.addMergedRegion(new CellRangeAddress(2,2,20, 23) );
        sheet.addMergedRegion(new CellRangeAddress(2,2,24, 27) );
        sheet.addMergedRegion(new CellRangeAddress(2,2,28, 31) );
        rowx.createCell(3).setCellValue("Обоснование ");
        rowx.getCell(3).setCellStyle(slt[1]);
        rowx.createCell(5).setCellValue("Физ. Объемы ");
        rowx.getCell(5).setCellStyle(slt[1]);
        rowx.createCell(11).setCellValue("Стоимость работ на период ТП, тыс.руб. (в тек.ур.цен.)");
        rowx.getCell(11).setCellStyle(slt[1]);
        rowx.createCell(14).setCellValue("Стоимость работ на 20.07");
        rowx.getCell(14).setCellStyle(slt[1]);
        rowx.createCell(16).setCellValue("1 неделя ТП ");
        rowx.getCell(16).setCellStyle(slt[1]);
        rowx.createCell(20).setCellValue("2 неделя ТП ");
        rowx.getCell(20).setCellStyle(slt[1]);
        rowx.createCell(24).setCellValue("3 неделя ТП ");
        rowx.getCell(24).setCellStyle(slt[1]);
        rowx.createCell(28).setCellValue("4 неделя ТП ");
        rowx.getCell(28).setCellStyle(slt[1]);

        rowx = sheet.createRow(rowi++);//создаем строку
        sheet.addMergedRegion(new CellRangeAddress(3,3,8, 10) );
        sheet.addMergedRegion(new CellRangeAddress(3,3,16, 17) );
        sheet.addMergedRegion(new CellRangeAddress(3,3,18, 19) );
        sheet.addMergedRegion(new CellRangeAddress(3,3,20, 21) );
        sheet.addMergedRegion(new CellRangeAddress(3,3,22, 23) );
        sheet.addMergedRegion(new CellRangeAddress(3,3,24, 25) );
        sheet.addMergedRegion(new CellRangeAddress(3,3,26, 27) );
        sheet.addMergedRegion(new CellRangeAddress(3,3,28, 29) );
        sheet.addMergedRegion(new CellRangeAddress(3,3,30, 31) );
        rowx.createCell(8).setCellValue("На период ТП");
        rowx.getCell(8).setCellStyle(slt[1]);
        rowx.createCell(16).setCellValue("Физ. Объем");
        rowx.getCell(16).setCellStyle(slt[1]);
        rowx.createCell(18).setCellValue("Стоимость работ");
        rowx.getCell(18).setCellStyle(slt[1]);
        rowx.createCell(20).setCellValue("Физ. Объем");
        rowx.getCell(20).setCellStyle(slt[1]);
        rowx.createCell(22).setCellValue("Стоимость работ");
        rowx.getCell(22).setCellStyle(slt[1]);
        rowx.createCell(24).setCellValue("Физ. Объем");
        rowx.getCell(24).setCellStyle(slt[1]);
        rowx.createCell(26).setCellValue("Стоимость работ");
        rowx.getCell(26).setCellStyle(slt[1]);
        rowx.createCell(28).setCellValue("Физ. Объем");
        rowx.getCell(28).setCellStyle(slt[1]);
        rowx.createCell(30).setCellValue("Стоимость работ");
        rowx.getCell(30).setCellStyle(slt[1]);
        rowx = sheet.createRow(rowi++);//создаем строку
        prnq(" *5*************************************** ");
        for (Elcol i: lecol) {//формирование шапки таблицы
            sheet.getRow(i.getNrow()).createCell(i.getNcol()).setCellValue(i.getName());
            sheet.getRow(i.getNrow()).getCell(i.getNcol()).setCellStyle(slt[1]);
 //           prnq(" "+i.getNcol()+"  "+i.getName()+" ^"+i.getNdt()+" "+i.getFrml());
        }//for wtu
        sheet.addMergedRegion(new CellRangeAddress(2,4,0, 0) );
        sheet.addMergedRegion(new CellRangeAddress(2,4,1, 1) );
        sheet.addMergedRegion(new CellRangeAddress(2,4,2, 2) );
        sheet.addMergedRegion(new CellRangeAddress(3,4,3, 3) );
        sheet.addMergedRegion(new CellRangeAddress(3,4,4, 4) );
        sheet.addMergedRegion(new CellRangeAddress(3,4,5, 5) );
        sheet.addMergedRegion(new CellRangeAddress(3,4,6, 6) );
        sheet.addMergedRegion(new CellRangeAddress(3,4,7, 7) );
        sheet.addMergedRegion(new CellRangeAddress(2,4,32, 32) );
        prnq(" *6* начинаю формировать рабочее поле таблицы * ");
        int rowbeg = rowi;
        for (Elmnt itn : elst){
            rowx = sheet.createRow(rowi);//создаем строку

            if (itn.getIdr()>=0) {
                for (Elcol i: lecol) {
                    int j=i.getNdt();//номер колонки данных
                    if (j==0) {
                        if (i.getFrml(rowi).length()<2)prnq(rowi+") ОШИБКА в колонке"+i.getNcol()+": "+i.getFrml(rowi));
                        else {
                            String f=i.getFrml(rowi);
//                            f="E3";
                            rowx.createCell(i.getNcol()).setCellFormula(f);
//                            rowx.getCell(i.getNcol()).setCellType(CellType.FORMULA);
                            rowx.getCell(i.getNcol()).setCellStyle(slt[i.getStl()]);
                        }
                    }else {
                        if (itn.getVal(i.getNdt()) instanceof String) {//является ли объект строкой
                            rowx.createCell(i.getNcol()).setCellValue((String) itn.getVal(i.getNdt()));
                            rowx.getCell(i.getNcol()).setCellStyle(slt[i.getStl()]);
                        } else if (itn.getVal(i.getNdt()) instanceof Integer) {//является ли объект числом
                            rowx.createCell(i.getNcol()).setCellValue((Integer) itn.getVal(i.getNdt()));
                            rowx.getCell(i.getNcol()).setCellStyle(slt[i.getStl()]);
                        } else if (itn.getVal(i.getNdt()) instanceof Double) {//является ли объект числом
                            rowx.createCell(i.getNcol()).setCellValue((Double) itn.getVal(i.getNdt()));
                            rowx.getCell(i.getNcol()).setCellStyle(slt[i.getStl()]);
                        }
                    }
                }//for wtu
            } else {
                cell = rowx.createCell(0);
                cell.setCellStyle(slt[2]);//установка стиля на ячейку
                cell.setCellType(CellType.STRING);
                cell.setCellValue(itn.getName());
                sheet.addMergedRegion(new CellRangeAddress(rowi,rowi,0, 7) );
            }//if
            rowi++;
        }//for elst (row)
        assert prnq(" *7* формирование нижней таблички* ");
        int rowend = (rowi++)-2;

        rowx = sheet.createRow(rowi++);//создаем строку
        rowx.createCell(2).setCellValue("Наименование СП организации");
        rowx.createCell(3).setCellValue("План на месяц ");
        rowx.createCell(4).setCellValue("План на дату отчета,");
        rowx.createCell(5).setCellValue("Факт  на дату отчета");
        rowx.createCell(6).setCellValue("Отклонение ");
        rowx.createCell(7).setCellValue("Численность ");
        for (int i=2; i<8; i++ ){
            rowx.getCell(i).setCellStyle(slt[1]);
        }
        String formula;

        for (String s: lpipl) {
 //           prnq(" "+s);
            rowx = sheet.createRow(rowi++);//создаем строку
            rowx.createCell(1).setCellValue(s);
            rowx.getCell(1).setCellStyle(slt[6]);

            rowx.createCell(2).setCellValue(s);
            rowx.getCell(2).setCellStyle(slt[3]);
            String f="COUNTIF(B$10:B$"+rowend+",B"+rowi+")";//ВНИМАНИЕ вместо ";" нужно писать ","
            rowx.createCell(0).setCellFormula(f);
            rowx.getCell(0).setCellStyle(slt[4]);
            formula = rowx.getCell(0).getCellFormula();
            evaluator = wbook.getCreationHelper().createFormulaEvaluator();
            cellValue = evaluator.evaluate(rowx.getCell(0));
            value = cellValue.getNumberValue();
            rowx.getCell(0).setCellValue(value);

            f="SUMIF(B$10:B$"+rowend+",B"+rowi+",L$10:L$"+rowend+")";
            int l=3;
            rowx.createCell(l).setCellFormula(f);
            rowx.getCell(l).setCellStyle(slt[5]);
            formula = rowx.getCell(l).getCellFormula();
            evaluator = wbook.getCreationHelper().createFormulaEvaluator();
            cellValue = evaluator.evaluate(rowx.getCell(l));
            value = cellValue.getNumberValue();
            rowx.getCell(l).setCellValue(value);

            l=4;
            f="SUMIF(B$10:B$"+rowend+",B"+rowi+",O$10:O$"+rowend+")";
            rowx.createCell(l).setCellFormula(f);
            rowx.getCell(l).setCellStyle(slt[5]);
            formula = rowx.getCell(l).getCellFormula();
            evaluator = wbook.getCreationHelper().createFormulaEvaluator();
            cellValue = evaluator.evaluate(rowx.getCell(l));
            value = cellValue.getNumberValue();
            rowx.getCell(l).setCellValue(value);

            l=5;
            f="SUMIF(B$10:B$"+rowend+",B"+rowi+",M$10:M$"+rowend+")";
            rowx.createCell(l).setCellFormula(f);
            rowx.getCell(l).setCellStyle(slt[5]);
            formula = rowx.getCell(l).getCellFormula();
            evaluator = wbook.getCreationHelper().createFormulaEvaluator();
            cellValue = evaluator.evaluate(rowx.getCell(l));
            value = cellValue.getNumberValue();
            rowx.getCell(l).setCellValue(value);

            l=6;
            f="E"+rowi+"-F"+rowi;
            rowx.createCell(l).setCellFormula(f);
            rowx.getCell(l).setCellStyle(slt[5]);
            formula = rowx.getCell(l).getCellFormula();
            evaluator = wbook.getCreationHelper().createFormulaEvaluator();
            cellValue = evaluator.evaluate(rowx.getCell(l));
            value = cellValue.getNumberValue();
            rowx.getCell(l).setCellValue(value);
        }
//-----------------------------------------------------------------------
        BgFile.savefl(wbook,filenameoutv);//сохранение результата работы
        assert prnq(" *Z* Main End ****");
    }//main
}//class
