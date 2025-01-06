/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package compileexcel;

import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;

/**
 *
 * @author matti
 */
public class GestoreExcel {
    static int contaUtenti(HSSFSheet  sheetD) throws IOException{
        int j=1;
        while(sheetD.getRow(j)!= null){
            if(new DatiRiga(j,sheetD).ruolo.equals("Pr"))
                j++;
        }
        return j;
    }
    static  HSSFSheet creaRighe(HSSFSheet  sheetK,HSSFSheet  sheetD) throws IOException {
        for(int i=6;i<contaUtenti(sheetD)*6;i++){
            sheetK.createRow(i);
        }
        return sheetK;
    }
}
