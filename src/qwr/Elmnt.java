package qwr;

import static qwr.Main.prnt;

public class Elmnt {
    private int idr;    //тип записи и номер записи -1=группирующий уровень
    private String  id;     //идентификатор
    private String  name;   //наименование
    private String  smin;   //инвентарный номер сметы
    private String  rd;     //раб.док
    private String  edi;    //единица измерния
    private double  ksm;    //общее количество
    private double  kost;   //остаток на начало
    private double  kpl;    //план на период
    private double  csm;    //общая стоимость
    private double  cst;    //остаток на начало
    private double  cpl;    //пллан стоимости на период
    private double  cpt;    //пллан стоимости на период в текущих
    private String  smln;   //локальный номер сметы
    private String  pipl;   //исполнитель
    private int     tst;    //-1 пропустить строку/0 оставиь/+1 включить  (Y)
    //--------------------------------------------
    public Object getVal(int i){
        switch (i){
            case 0 :return " ";
            case 1 :return this.id;   //идентификатор
            case 2 :return this.name; //наименование
            case 3 :return this.rd;   //раб.док
            case 4 :return this.smin; //инвентарный номер сметы
            case 5 :return this.smln; //локальный номер сметы
            case 6 :return this.pipl; //исполнитель
            case 7 :return this.edi;  //единица измерния
            case 8 :return this.ksm;  //общее количество
            case 9 :return this.kost; //остаток на начало
            case 10:return this.kpl;  //план на период
            case 11:return this.csm;  //общая стоимость
            case 12:return this.cst;  //остаток на начало
            case 13:return this.cpl;  //пллан стоимости на период
            case 14:return this.idr;  //тип записи и номер записи -1=группирующий
            case 15:return (double)0;
            case 16:return " ";
            case 17:return this.cpt;    //пллан стоимости на период в текущих
            default:return " @ "; //не пропустить строку
        }//switch
    }//getVal
    public Elmnt(){
        this.idr = 0;
        this.id = "";
        this.name = "";
        this.smin = "";
        this.rd = "";
        this.edi = "";
        this.ksm = 0;
        this.kost = 0;
        this.kpl = 0;
        this.csm = 0;
        this.cst = 0;
        this.cpl = 0;
        this.cpt = 0;
        this.smln = "";
        this.pipl ="";
        this.tst = 0;//-1 пропустить строку/0 оставиь/+1 включить
    }//Elmnt
    public Elmnt(Elmnt q){//копирование элемента
        this.idr = q.idr;
        this.id = q.id;
        this.name = q.name;
        this.smin = q.smin;
        this.rd = q.rd;
        this.edi = q.edi;
        this.ksm = q.ksm;
        this.kost = q.kost;
        this.kpl = q.kpl;
        this.csm = q.csm;
        this.cst = q.cst;
        this.cpl = q.cpl;
        this.cpt = q.cpt;
        this.smln = q.smln;
        this.pipl =q.pipl;
        this.tst = q.tst;//-1 пропустить строку/0 оставиь/+1 включить
    }//Elmnt
    public void clear(){
        this.idr = 0;
        this.id = "";
        this.name = "";
        this.smin = "";
        this.rd = "";
        this.edi = "";
        this.ksm = 0;
        this.kost = 0;
        this.kpl = 0;
        this.csm = 0;
        this.cst = 0;
        this.cpl = 0;
        this.cpt = 0;
        this.smln = "";
        this.pipl ="";
        this.tst = 0;//-1 пропустить строку/0 оставиь/+1 включить
    }//clear
    public int      getIdr() {return idr; }
    public String   getName() { return name; }
    public String   getPipl() { return pipl; }//исполнитель
    public int getTst() { return tst; }

    public boolean isAdd(){//есть плановые назначения и есть наименование
        if (this.tst<0){  prnt("--- Исключаю строку --- "); return false;}
        boolean b=((idr==-1)&&!(this.name.equals("")) )||
                (this.kpl + this.cpl != 0) && (idr>=0);
        if (!b) prnt("-+- Пропускаю строку -+- idr:"+idr+" E:"+
                (this.kpl + this.cpl)+" --- "+this.name);
        if (this.tst>0){  prnt("-++ Включаю строку ++- "); return true;}
        return  b;
    }//isAdd

    public void setstr(int i,String s){//Читаю строку и разбрасываю
        //очищаю строку от лишних символов
        s=s.trim();
        int length = s.length();
        char[] y = new char[length+1];
        int newLen = 0;
        y[newLen]= s.charAt(newLen++);
        for (int  j = newLen ; j < length; j++) {
            char ch=s.charAt(j);
            if ((ch > 32) || (ch==32 && y[newLen-1]>32)) {
                y[newLen] = ch;
                newLen++;
            }
        }
         String x = new String(y, 0, newLen);
        //применяю очищенную строку
        switch (i){
        case 0: this.id=x; break;   //идентификатор (A)
        case 5: this.name =x;       //наименование  (F)
                this.idr=-1;        //группировка
                break;
        case 6: if (this.idr==0 && this.name.equals("")){//(G)
                    this.idr=-1;     //группировка
                    this.name =x;    //наименование
                } else this.rd=x;   //раб.док
                break;//
        case 7: this.smin =x;       //инвентарный номер сметы   (H)
                if (this.idr<0) this.idr=0;
                break;
        case 8: this.smln =x; //локальный номер сметы   (I)
                if (this.idr<0) this.idr=0;
                break;
        case 9: this.name =x; //наименование    (J)
                if (this.idr<0) this.idr=0;
                break;
        case 11: this.edi=x; break;  //единица измерния (L)
        case 12:                     //общее количество (M)
        case 13:                     //остаток на начало (N)
        case 14:                     //план на период (O)
        case 15:                     //общая стоимость (P)
        case 16:                                // (Q)
        case 17:                     //остаток на начало (R)
        case 18:                     //пллан стоимости на период (S)
            break;
        case 21: this.pipl=x; break; //исполнитель
        default:prnt("&~");
    }}//setstr
    public void setdbl(int i,double x){ switch (i){//читаю число и разбрасываю
        case 5: this.idr= (int) x; break;  //номер записи
        case 12: this.ksm=x; break;  //общее количество
        case 13: this.kost=x; break; //остаток на начало
        case 14: this.kpl=x; break;  //план на период
        case 15: this.csm=x; break;  //общая стоимость
        case 17: this.cst=x; break;  //остаток на начало
        case 18: this.cpl=x; break;  //пллан стоимости на период  (S)
        case 19: this.cpt=x; break;  //пллан стоимости на период  (T)
        case 24: this.tst=(int) x; prnt("W"+x); break;//-1 пропустить/0 оставиь/+1 включить
        default:  prnt("$~");//-1 пропустить/0 оставиь/+1 включить
    }}//setdbl
    public String prn(){ return ""+ this.idr+":"+ this.id+":"+ this.name+":\n"+
        this.rd+" : "+ this.smin +" : "+ this.smln+" : "+ this.edi+" :\n "+
        this.ksm+" : "+ this.kost+" : "+ this.kpl+" : "+ this.csm+" : "+
        this.cst+" : "+ this.cpl+" :~"+ this.pipl+" ;";
    }
}//class