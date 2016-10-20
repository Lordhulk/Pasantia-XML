
import java.awt.HeadlessException;
import jxl.*;
import jxl.write.*;
import jxl.Workbook;
import java.io.File;
import jxl.Cell;
import javax.swing.*;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import jxl.read.biff.BiffException;

public class NuevoArchivo {
  public WritableWorkbook hojaexcel2;
  public WritableSheet hojanueva1;
  public String nombre_del_excel2;
  
   public void NuevoArchivo2() {
      try {
             
     String nombre_del_excel = JOptionPane.showInputDialog("Ingrese el nombre del archivo donde quiere guardar los datos sacados"); 
     String formato;
     formato = ".xls";
     nombre_del_excel2 = nombre_del_excel + formato; 
     String datosmuestra = "datosmuestras";
           WritableWorkbook nuevoexcel = Workbook.createWorkbook(new File(nombre_del_excel2));
          WritableSheet hojanueva = nuevoexcel.createSheet(datosmuestra,1);
       //WritableSheet hojanueva = workbook.("datos muestra", 1);  
          
          
 this.hojanueva1 = hojanueva;
this.hojaexcel2 = nuevoexcel;
//nuevoexcel.close();
  //  nuevoexcel.write();    
 
   } catch (Exception ioe) {
            ioe.printStackTrace();
        }
        
   }
   public String getnombre_archivo(){
       return nombre_del_excel2;
       
   }
   public WritableSheet getnombre_archivo2(){
       return hojanueva1;
       
   }
      public WritableWorkbook getnombre_archivo3(){
       return hojaexcel2;
       
   }
}