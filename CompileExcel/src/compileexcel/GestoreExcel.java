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
    static int contaUtenti(HSSFSheet  sheetD,boolean PrCo) throws IOException{
        int j=1;
        int i=0;
        while(sheetD.getRow(j)!= null){
            
            if(PrCo){
                if(sheetD.getRow(j).getCell((short)7).getStringCellValue().equals("Pr")){
                    i++;
                    j++;}
                else
                    j++;
            }else{
                i++;
                j++;}
        }
        return i;
    }
    static  HSSFSheet creaRighe(HSSFSheet  sheetK,HSSFSheet  sheetD) throws IOException {
        for(int i=6;i<contaUtenti(sheetD,true)*6;i++){
            sheetK.createRow(i);
        }
        return sheetK;
    }
}
