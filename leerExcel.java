
import jxl.*;
import jxl.write.*;
import jxl.Workbook;
import java.io.File;
import jxl.Cell;
import javax.swing.*;
import java.io.BufferedReader;
import java.io.InputStreamReader;

 
public class leerExcel {

int g2;

    public void leerArchivoExcel(String archivoOrigen,int g1,String archivoDestino,WritableSheet nombre_hoja,WritableWorkbook hojaexcel) {
        try {
            Workbook archivoExcel = Workbook.getWorkbook(new File(archivoOrigen));
        Sheet hoja = archivoExcel.getSheet(1);
        BufferedReader in = new BufferedReader(new InputStreamReader(System.in));
            String celda1;

            this.g2 = g1 + g2;
            //JOptionPane.showMessageDialog(null,g1);
    celda1 = JOptionPane.showInputDialog("Ingrese el numero de TER: ");
            int celda2 = Integer.parseInt(celda1);   
            String a3[] = new String [10];
            int numeros[]= new int [10];


        for (int Y=0;Y!=7;Y++)
            {
        
         
         Cell a2= hoja.getCell(Y,celda2);
        String stringa2 = a2.getContents();
        a3 [Y] = stringa2;

        switch(Y){
            case 1:
         String string_a2[] = stringa2.split ("C");
         numeros[0] = Integer.parseInt(string_a2[1]);
                break;
            case 2:
         String string_a3[] = stringa2.split ("M");
         numeros[1] = Integer.parseInt(string_a3[1]);
                break;
                            case 3:
         String string_a4[] = stringa2.split ("T");
         numeros[2] = Integer.parseInt(string_a4[1]);
                break;
                            case 4:
         String string_a5[] = stringa2.split ("R");
         numeros[3] = Integer.parseInt(string_a5[1]);
         break;
         default:
        }
            }
        

       Sheet hoja2 = archivoExcel.getSheet(5);
       Cell lec1 = hoja2.getCell(3,numeros[0]+6);
       String lec_1 = lec1.getContents();
       Cell lec2 = hoja2.getCell(4,numeros[0]+6);
       String lec_2 = lec2.getContents();
       Cell lec3 = hoja2.getCell(5,numeros[0]+6);
       String lec_3 = lec3.getContents();
       
       
       Sheet hoja3 = archivoExcel.getSheet(4);
       Cell lem1 = hoja3.getCell(3,numeros[1]+6);
       String lem_1 = lem1.getContents();
       Cell lem2 = hoja3.getCell(4,numeros[1]+6);
       String lem_2 = lem2.getContents();
       Cell lem3 = hoja3.getCell(5,numeros[1]+6);
       String lem_3 = lem3.getContents();
       
       
       Sheet hoja4 = archivoExcel.getSheet(7);
       Cell let1 = hoja4.getCell(3,numeros[2]+6);
       String let_1 = let1.getContents();
       Cell let2 = hoja4.getCell(4,numeros[2]+6);
       String let_2 = let2.getContents();
       Cell let3 = hoja4.getCell(5,numeros[2]+6);
       String let_3 = let3.getContents();
       
       
       Sheet hoja5 = archivoExcel.getSheet(7);
       Cell ler1 = hoja5.getCell(3,numeros[3]+6);
       String ler_1 = ler1.getContents();
       Cell ler2 = hoja5.getCell(4,numeros[3]+6);
       String ler_2 = ler2.getContents();
       Cell ler3 = hoja5.getCell(5,numeros[3]+6);
       String ler_3 = ler3.getContents();     
        archivoExcel.close();   
      // WritableSheet nombre_hoja2 = hojaexcel.createSheet("datos muestra", 1);  
        
        
        
      //JOptionPane.showMessageDialog(null,nombre_hoja);
       nombre_hoja.addCell(new jxl.write.Label(1,g1,a3[1]));
       nombre_hoja.addCell(new jxl.write.Label(2,g1,lec_1));
       nombre_hoja.addCell(new jxl.write.Label(4,g1,lec_2));
       nombre_hoja.addCell(new jxl.write.Label(6,g1,lec_3));
 
       nombre_hoja.addCell(new jxl.write.Label(1,g1+1,a3[2]));
       nombre_hoja.addCell(new jxl.write.Label(2,g1+1,lem_1));
       nombre_hoja.addCell(new jxl.write.Label(4,g1+1,lem_2));
       nombre_hoja.addCell(new jxl.write.Label(6,g1+1,lem_3));
 
       nombre_hoja.addCell(new jxl.write.Label(1,g1+2,a3[3]));
       nombre_hoja.addCell(new jxl.write.Label(2,g1+2,let_1));
       nombre_hoja.addCell(new jxl.write.Label(4,g1+2,let_2));
       nombre_hoja.addCell(new jxl.write.Label(6,g1+2,let_3));
 
      nombre_hoja.addCell(new jxl.write.Label(1,g1+3,a3[4]));
      nombre_hoja.addCell(new jxl.write.Label(2,g1+3,ler_1));
      nombre_hoja.addCell(new jxl.write.Label(4,g1+3,ler_2));
      nombre_hoja.addCell(new jxl.write.Label(6,g1+3,ler_3));   



//hojaexcel.write();    
//hojaexcel.close();

        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
 
    }
public int contador(int ga){
String soption;
int opcion;
soption = JOptionPane.showInputDialog("Desea buscar otro TER? 1:si 2:no ");
opcion=Integer.parseInt(soption);
//archivoExcel_2.write(); 
switch (opcion)
{
case 1:
ga=ga+5;
this.g2 = ga;
//opcion=2;
break; 
case 2:
//hojaexcel.close();
ga=0;

break; 
default:
JOptionPane.showMessageDialog(null,"Para salir verifique la opción deseada");

}
return ga;

}


}