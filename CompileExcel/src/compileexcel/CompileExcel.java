/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package compileexcel;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

/**
 *
 * @author matti
 */
public class CompileExcel {

    /**
     * @param args the command line arguments
     */
     public static void main(String[] args) throws FileNotFoundException, IOException {
        Scanner scan=new Scanner(System.in);
        System.out.println("scrivi il percorso del file sorgente");
        String percorsoKonta =scan.next();
        System.out.println("scrivi il percorso del file destinazione");
        String percorsoDanea =scan.next();
        CompilaFile ciao=new CompilaFile(percorsoDanea,percorsoKonta);
        ciao.inserisciDati();
    }
    
}
