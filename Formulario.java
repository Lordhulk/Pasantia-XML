import java.awt.TextField;
import javax.swing.*;
import java.awt.event.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.write.*;
public class Formulario extends JFrame implements ActionListener{
    public JButton boton1,boton2,boton3;
    leerExcel excel1 = new leerExcel();

    
    
    public Formulario() {
        setLayout(null);

        boton1=new JButton("iniciar programa");
        boton1.setBounds(10,100,190,30);
        add(boton1);
        boton1.addActionListener(this);
        boton2=new JButton("cerrar programa");
        boton2.setBounds(210,100,190,30);
        add(boton2);
        boton2.addActionListener(this);
        boton3=new JButton("¿buscar otro?");
        boton3.setBounds(210,100,90,30);
        add(boton3);
        boton3.addActionListener(this);  	
    }
    public void actionPerformed(ActionEvent e) {
        if (e.getSource()==boton1) {

            setTitle("PROGRAMA EN USO PORFAVOR ESPERE...");
            Ventanaexcel ventana = new Ventanaexcel();
            ventana.actionPerformed(e);
            try {
                Thread.sleep(2000);
            } catch (InterruptedException ex) {
                Logger.getLogger(Formulario.class.getName()).log(Level.SEVERE, null, ex);
            }
            setTitle("GRACIAS POR ESPERAR YA PUEDE CERRAR EL PROGRAMA");

            
           
        }
         if (e.getSource()==boton2) 
             
 
System.exit(0);

  }
    
   
}