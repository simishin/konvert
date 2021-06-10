package qwr.utl;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

public class BgFile {
public static String flin  ="perq.xlsx";//имя файла результата
public static String flout;  //имя файла результата
private static final String mask ="tempPlan";
public static boolean prnq(String s){System.out.println(s); return true;}
public static boolean prnt(String s){System.out.print(s); return true;}
//----------------------------------------------------------------------------------
public static boolean comstring(String[] args) {//разбор командной строки

    boolean tstAssert = false;
    assert tstAssert = true;
    if (tstAssert) { System.out.println("!!!       Assertions are enabled");
    } else { System.out.println("!!!     Assertions are disabled. set (-ea)"); }

    if (args.length!=0) { //аргументы есть
        for (String s:args){
            assert prnq("~"+s);
            if(s.startsWith("-h")||s.startsWith("-H")||args.length>1){
                prnq("Программа формирования формул и группировок в темплане.");
                prnq("Один аргумент - имя исходного файла XLSX");
                System.exit(0);
            }//if

            if (!s.endsWith(".xlsx")) continue; //это не файл  XLSX.
            File item = new File(s);//существует внутри цикла
            if (item.exists()) {
                flin = s;
                break;//файл существует конец просмотра аргументов
            }
        }//for args
    }
    //проверяю существование файла
    File item = new File(flin);//существует внутри цикла
    if (!item.isFile()) {
        prnq("Файл с именем "+ flin +" не найден. Выполнение программы завершено.");
        System.exit(1);
    }//if
    //создаю имя файла результата
//    String s= flin.replaceFirst(".",mask.concat("_"));//результат tempPlan_021-01.xlsx
    flout= flin.replaceFirst(".x",mask.concat(".x"));//результат 2021-01tempPlan.xlsx
//    prnq("Читаю файл "+flin);
    return true;
}//comstring
    //--------------------------------------------------------------------------
public static void savefl(XSSFWorkbook wbook){//сохранение результата работы
    prnq("Сохранение результата работы в файле "+ BgFile.flout);
    try  {
        FileOutputStream out = new FileOutputStream(new File(BgFile.flout));
        wbook.write(out);
        out.close();
        System.out.println( "Результат сохранен");
    }
    catch (Exception e){ e.printStackTrace(); }
}//savefl
    public static void savefl(XSSFWorkbook wbook, String file){//сохранение результата работы
        prnq("Сохранение результата работы в файле "+ file);
        try  {
            FileOutputStream out = new FileOutputStream(new File(file));
            wbook.write(out);
            out.close();
            System.out.println( "Результат сохранен");
        }
        catch (Exception e){ e.printStackTrace(); }
    }//savefl
//----------------------------------------------------------------------------------
}//class srtd
