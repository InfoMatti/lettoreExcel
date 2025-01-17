/*
* Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
* Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
*/
package compileexcel;
import java.io.FileNotFoundException;
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
    
    HSSFSheet  sheet;
    
    public LetturaFileDanea(HSSFSheet  sheet) throws FileNotFoundException, IOException{
        this.sheet=sheet;
    }
    
    ArrayList<String> getRuolo() throws IOException{
        ArrayList<String> ruolo=new ArrayList<String>();
        for(int i=1;i<GestoreExcel.contaUtenti(sheet,false)+1;i++){
            ruolo.add(sheet.getRow(i).getCell((short)7).getStringCellValue());
            // System.out.println( numAlloggio.get(i-1));
        }
        return ruolo;
    }
    
    ArrayList<String> getPiano() throws IOException{
        ArrayList<String> piano=new ArrayList<String>();
        for(int i=1;i<GestoreExcel.contaUtenti(sheet,false)+1;i++){
            piano.add(sheet.getRow(i).getCell((short)3).getStringCellValue());
            // System.out.println( numAlloggio.get(i-1));
        }
        return piano;
    }
    
    ArrayList<String> getScalaNumAlloggio() throws IOException{
        ArrayList<String> numAlloggio=new ArrayList<String>();
        for(int i=1;i<GestoreExcel.contaUtenti(sheet,false)+1;i++){
            numAlloggio.add(getScala().get(i-1)+getNumApp().get(i-1));
            // System.out.println( numAlloggio.get(i-1));
        }
        return numAlloggio;
    }
    
    ArrayList<String> getNumApp() throws IOException{
        ArrayList<String> numApp=new ArrayList<String>();
        for(int i=1;i<GestoreExcel.contaUtenti(sheet,false)+1;i++){
            numApp.add(sheet.getRow(i).getCell((short)2).getStringCellValue());
        }
        return numApp;
    }
    
    ArrayList<String> getScala() throws IOException{
        ArrayList<String> scala=new ArrayList<>();
        for(int i=1;i<GestoreExcel.contaUtenti(sheet,false)+1;i++){
            scala.add(sheet.getRow(i).getCell((short)1).getStringCellValue());
        }
        return scala;
    }
    
    ArrayList<String> getNome() throws IOException{
        ArrayList<String> nomeProprietario=new ArrayList<String>();
        for(int i=1;i<GestoreExcel.contaUtenti(sheet,false)+1;i++){
            nomeProprietario.add(sheet.getRow(i).getCell((short)9).getStringCellValue());
            //System.out.println( nomeUtente.get(i-1));
        }
        return nomeProprietario;
    }
    
    ArrayList<String> getCF() throws IOException{
        ArrayList<String> CF=new ArrayList<String>();
        for(int i=1;i<GestoreExcel.contaUtenti(sheet,false)+1;i++){
            CF.add(sheet.getRow(i).getCell((short)22).getStringCellValue());
            //System.out.println( CF.get(CF.size()-1));
        }
        return CF;
    }
    
    ArrayList<String> getNumTell() throws IOException{
        ArrayList<String> numTell=new ArrayList<String>();
        for(int i=1;i<GestoreExcel.contaUtenti(sheet,false)+1;i++){
            numTell.add(sheet.getRow(i).getCell((short)24).getStringCellValue());
            //System.out.println( numTell.get(numTell.size()-1));
        }
        return numTell;
    }
    
    ArrayList<String> getEmail() throws IOException{
        ArrayList<String> email=new ArrayList<String>();
        for(int i=1;i<GestoreExcel.contaUtenti(sheet,false)+1;i++){
            email.add(sheet.getRow(i).getCell((short)28).getStringCellValue());
            //System.out.println( email.get(email.size()-1));
        }
        return email;
    }
    
    ArrayList<String> getIndirizzo() throws IOException {
        ArrayList<String> indirizzo = new ArrayList<String>();
        for (int i = 1; i < GestoreExcel.contaUtenti(sheet,false)+1; i++) {
            String fullString = sheet.getRow(i).getCell((short)11).getStringCellValue();
            indirizzo.add(fullString.split("\\d", 2)[0]);
            if(indirizzo.get(indirizzo.size()-1).contains(",")){
                indirizzo.set(indirizzo.size()-1, indirizzo.get(indirizzo.size()-1).replace(",", ""));
            }
            //System.out.println(indirizzo.get(indirizzo.size() - 1));
        }
        return indirizzo;
    }
    
    ArrayList<String> getCivico() throws IOException {
        ArrayList<String> civico = new ArrayList<String>();
        Pattern pattern = Pattern.compile("\\d+");
        for (int i = 1; i < GestoreExcel.contaUtenti(sheet,false)+1; i++) {
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
