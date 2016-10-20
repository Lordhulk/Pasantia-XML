import javax.swing.JOptionPane;
import javax.swing.*;
import java.awt.event.*;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
public class Ventanaexcel extends JFrame implements ActionListener{
    public void actionPerformed(ActionEvent e) {
leerExcel excel = new leerExcel(); 
int contador = 0;
String nombrearchivo;
WritableSheet nombrearchivo2;
WritableWorkbook nombrearchivo3;
NuevoArchivo nuevo = new NuevoArchivo();
//NuevoArchivo nuevo2 = new NuevoArchivo();
nuevo.NuevoArchivo2();
//nuevo.NuevoArchivo2();
nombrearchivo = nuevo.getnombre_archivo();
nombrearchivo2 = nuevo.getnombre_archivo2();
nombrearchivo3 = nuevo.getnombre_archivo3();
do{  
excel.leerArchivoExcel("Recursos Informáticos de extensión.xls",contador,nombrearchivo,nombrearchivo2,nombrearchivo3);
contador = excel.contador(contador);
}while(contador!=0);
        
            
        try {
            nombrearchivo3.write();  
            nombrearchivo3.close();
        } catch (IOException ex) {
            Logger.getLogger(Ventanaexcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WriteException ex) {
            Logger.getLogger(Ventanaexcel.class.getName()).log(Level.SEVERE, null, ex);
        }


}
}
//System.exit(0)