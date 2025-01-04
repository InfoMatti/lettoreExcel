/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package compileexcel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author matti
 */
public class LetturaFileDanea {
    
        FileInputStream fileDanea;
        HSSFWorkbook fD;
        HSSFSheet  sheet;
        
        public LetturaFileDanea(String percorsoDanea) throws FileNotFoundException, IOException{
            this.fileDanea =new  FileInputStream(new File(percorsoDanea/*"D:\\LavoroPapa\\FileDanea.xls"*/));
            this.fD = new HSSFWorkbook(fileDanea);
            this.sheet = fD.getSheetAt(0);
        }
        int contaUtenti(){
            int j=1;
             while(sheet.getRow(j)!= null){
                 j++;
            }
            return j;
        }
         ArrayList<String> getRuolo() throws IOException{
           ArrayList<String> ruolo=new ArrayList<String>();     
           for(int i=1;i<contaUtenti();i++){
                    ruolo.add(sheet.getRow(i).getCell((short)7).getStringCellValue());
               // System.out.println( numAlloggio.get(i-1));
            }
            return ruolo;
        }
         
         ArrayList<String> getScalaNumAlloggio() throws IOException{
           ArrayList<String> numAlloggio=new ArrayList<String>();     
           for(int i=1;i<contaUtenti();i++){
                    numAlloggio.add(getScala().get(i-1)+getNumApp().get(i-1));
               // System.out.println( numAlloggio.get(i-1));
            }
            return numAlloggio;
        }
        
         ArrayList<String> getNumApp() throws IOException{
            ArrayList<String> numApp=new ArrayList<String>();
            for(int i=1;i<contaUtenti();i++){
                numApp.add(sheet.getRow(i).getCell((short)2).getStringCellValue());
            }
            return numApp;
        }
        
         ArrayList<String> getScala(){
            ArrayList<String> scala=new ArrayList<>();
            for(int i=1;i<contaUtenti();i++){
                scala.add(sheet.getRow(i).getCell((short)1).getStringCellValue());
            }
            return scala;
        }
        
         ArrayList<String> getNome(){
            ArrayList<String> nomeProprietario=new ArrayList<String>();
            for(int i=1;i<contaUtenti();i++){
                    nomeProprietario.add(sheet.getRow(i).getCell((short)9).getStringCellValue());
                //System.out.println( nomeUtente.get(i-1));
            }
            return nomeProprietario;
        }

         ArrayList<String> getCF(){
              ArrayList<String> CF=new ArrayList<String>();
              for(int i=1;i<contaUtenti();i++){
                    CF.add(sheet.getRow(i).getCell((short)22).getStringCellValue());  
                   //System.out.println( CF.get(CF.size()-1));
              }
              return CF;
        }
        
         ArrayList<String> getNumTell(){
              ArrayList<String> numTell=new ArrayList<String>();
              for(int i=1;i<contaUtenti();i++){
                    numTell.add(sheet.getRow(i).getCell((short)24).getStringCellValue());  
                    //System.out.println( numTell.get(numTell.size()-1));
              }
              return numTell;
        }
        
         ArrayList<String> getEmail(){
              ArrayList<String> email=new ArrayList<String>();
              for(int i=1;i<contaUtenti();i++){
                    email.add(sheet.getRow(i).getCell((short)28).getStringCellValue());  
                    //System.out.println( email.get(email.size()-1));
              }
              return email;
        }
        
        ArrayList<String> getIndirizzo() {
        ArrayList<String> indirizzo = new ArrayList<String>();
        for (int i = 1; i < contaUtenti(); i++) {
            String fullString = sheet.getRow(i).getCell((short)11).getStringCellValue();
            indirizzo.add(fullString.split("\\d", 2)[0]);
            if(indirizzo.get(indirizzo.size()-1).contains(",")){
                indirizzo.set(indirizzo.size()-1, indirizzo.get(indirizzo.size()-1).replace(",", ""));
            }
            //System.out.println(indirizzo.get(indirizzo.size() - 1));
        }
        return indirizzo;
    }
    
        ArrayList<String> getCivico() {
        ArrayList<String> civico = new ArrayList<String>();
        Pattern pattern = Pattern.compile("\\d+");
        for (int i = 1; i < contaUtenti(); i++) {
        String fullString = sheet.getRow(i).getCell((short)11).getStringCellValue();
        Matcher matcher = pattern.matcher(fullString);
        StringBuilder numeri = new StringBuilder();
        while (matcher.find()) {
            numeri.append(matcher.group());
        }
        civico.add(numeri.toString());
        //System.out.println(civico.get(civico.size() - 1));
    }
        return civico;
    }

}
