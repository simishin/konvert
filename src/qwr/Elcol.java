package qwr;
//класс описания колонок
public class Elcol {
    static int counter=0;
    private final int ncol;       //порядковый номер колонки
    private final int nrow;       //номр строки куда пишется шапка таблицы
    private final int wid;        //ширина колокки
    private final String name;   //наименование колонки в шапке
    private final int krow;     //порядковый номер нижней строки шапка таблицы
    private final int stl;       //стиль
    private final int ndt;        //номер колонки данных
    private final String frml;    //формула в ячейке
    public int getNrow() { return nrow; }//порядковый номер строки
    public int getNcol() { return ncol; }//порядковый номер колонки
    public int getWid() { return wid; }//ширина колокки
    public String getName() { return name; }//наименование колонки
    public int getKrow() { return krow; }//порядковый номер нижней строки
    public int getStl() { return stl; }//размер шрифта/стиль
    public int getNdt() { return ndt; }//номер колонки данных
    /** @param i номер строки для подстановки в формулу
     * @return преобразованная формула с измененными номерами строк */
    public String getFrml(int i) {
        if (this.frml.length()<2 || i<1) {  return "Error";  }
        StringBuilder s= new StringBuilder(this.frml);
        s.ensureCapacity(64);//увеличиваю объем буфера
        String d= Integer.toString(i+1);
        Boolean b=false;
        for (int k=s.length()-1, n=k; k>=0; k--){
            if (s.charAt(k)>='0'&& s.charAt(k)<='9') {
                if (!b) { b = true; n = k; } //это цифра
            }
            else {//не цифра
                if (s.charAt(k)!='$'&& b) { s.replace(k+1,n+1,d); }
                b=false;
            }
        }// s.delete(k+1,n+1);//s.insert(k+1,d);
        return String.valueOf(s);
    }//формула в ячейке
    public Elcol(int wid, String name, int nrow, int ndt, int stl) {
        this.ncol= counter++;//порядковый номер колонки
        this.nrow=nrow;
        this.wid = wid;//ширина колокки
        this.name = name;//наименование колонки
        this.krow = nrow;
        this.stl = stl;
        this.ndt = ndt;//номер колонки данных
        this.frml = "";//формула в ячейке
    }
    public Elcol(int wid, String name, int nrow, int ndt, int stl, String frml) {
        this.ncol= counter++;//порядковый номер колонки
        this.nrow=nrow;
        this.wid = wid;//ширина колокки
        this.name = name;//наименование колонки
        this.krow = nrow;
        this.stl = stl;
        this.ndt = 0;//номер колонки данных
        this.frml = frml;//формула в ячейке
    }
}//class