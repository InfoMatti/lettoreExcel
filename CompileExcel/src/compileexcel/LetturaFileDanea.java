/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package compileexcel;
import java.io.IOException;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.hssf.usermodel.HSSFSheet;

/**
 *
 * @author matti
 */
public class LetturaFileDanea {
       static int contaUtenti(HSSFSheet sheet ){
            int j=1;
             while(sheet.getRow(j)!= null){
                 j++;
            }
            return j-getNomeInquilino(sheet).size();
        }
        
        static ArrayList<String> getScalaNumAlloggio(HSSFSheet sheet) throws IOException{
           ArrayList<String> numAlloggio=new ArrayList<String>();     
           for(int i=1;i<contaUtenti(sheet);i++){
                    numAlloggio.add(getScala(sheet).get(i-1)+getNumApp(sheet).get(i-1));
               // System.out.println( numAlloggio.get(i-1));
            }
            return numAlloggio;
        }
        
        static ArrayList<String> getNumApp( HSSFSheet sheet) throws IOException{
            ArrayList<String> numApp=new ArrayList<String>();
            for(int i=1;i<contaUtenti(sheet);i++){
                numApp.add(sheet.getRow(i).getCell((short)2).getStringCellValue());
            }
            return numApp;
        }
        
        static ArrayList<String> getScala(HSSFSheet sheet){
            ArrayList<String> scala=new ArrayList<>();
            for(int i=1;i<contaUtenti(sheet);i++){
                scala.add(sheet.getRow(i).getCell((short)1).getStringCellValue());
            }
            return scala;
        }
        
        static ArrayList<String> getNomeProprietario(HSSFSheet sheet){
            ArrayList<String> nomeProprietario=new ArrayList<String>();
            for(int i=1;i<contaUtenti(sheet);i++){
                if(sheet.getRow(i).getCell((short)7).getStringCellValue().equals("Pr"))
                    nomeProprietario.add(sheet.getRow(i).getCell((short)9).getStringCellValue());
                //System.out.println( nomeUtente.get(i-1));
            }
            return nomeProprietario;
        }
        
        static ArrayList<String> getNomeInquilino(HSSFSheet sheet){
            ArrayList<String> nomeInquilino=new ArrayList<String>();
            for(int i=1;i<contaUtenti(sheet);i++){
                if(sheet.getRow(i).getCell((short)7).getStringCellValue().equals("Pr")==false)
                    nomeInquilino.add(sheet.getRow(i).getCell((short)9).getStringCellValue());
                //System.out.println( nomeUtente.get(i-1));
            }
            return nomeInquilino;
        }
        static ArrayList<String> getCF(HSSFSheet sheet){
              ArrayList<String> CF=new ArrayList<String>();
              for(int i=1;i<contaUtenti(sheet);i++){
                    CF.add(sheet.getRow(i).getCell((short)22).getStringCellValue());  
                   //System.out.println( CF.get(CF.size()-1));
              }
              return CF;
        }
        
        static ArrayList<String> getNumTell(HSSFSheet sheet){
              ArrayList<String> numTell=new ArrayList<String>();
              for(int i=1;i<contaUtenti(sheet);i++){
                    numTell.add(sheet.getRow(i).getCell((short)24).getStringCellValue());  
                    //System.out.println( numTell.get(numTell.size()-1));
              }
              return numTell;
        }
        
        static ArrayList<String> getEmail(HSSFSheet sheet){
              ArrayList<String> email=new ArrayList<String>();
              for(int i=1;i<contaUtenti(sheet);i++){
                    email.add(sheet.getRow(i).getCell((short)28).getStringCellValue());  
                    //System.out.println( email.get(email.size()-1));
              }
              return email;
        }
        
    static ArrayList<String> getIndirizzo(HSSFSheet sheet) {
        ArrayList<String> indirizzo = new ArrayList<String>();
        for (int i = 1; i < contaUtenti(sheet); i++) {
            String fullString = sheet.getRow(i).getCell((short)11).getStringCellValue();
            indirizzo.add(fullString.split("\\d", 2)[0]);
            if(indirizzo.get(indirizzo.size()-1).contains(",")){
                indirizzo.set(indirizzo.size()-1, indirizzo.get(indirizzo.size()-1).replace(",", ""));
            }
            //System.out.println(indirizzo.get(indirizzo.size() - 1));
        }
        return indirizzo;
    }
    
    static ArrayList<String> getCivico(HSSFSheet sheet) {
        ArrayList<String> civico = new ArrayList<String>();
        Pattern pattern = Pattern.compile("\\d+");
        for (int i = 1; i < contaUtenti(sheet); i++) {
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