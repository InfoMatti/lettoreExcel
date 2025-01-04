/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package compileexcel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
/**
 *
 * @author matti
 */
class CompilaFile {
     FileInputStream fileDanea;
    HSSFWorkbook fD;
    FileInputStream fileKonta;
    HSSFWorkbook fK;
    HSSFSheet sheetD;
    HSSFSheet sheetK;
    String percorsoKonta;
        
    public CompilaFile(String percorsoKonta,String percorsoDanea) throws FileNotFoundException, IOException{
        this.percorsoKonta=percorsoKonta;
        this.fileKonta = new  FileInputStream(new File(percorsoKonta/*"D:\\LavoroPapa\\ImportUnitàImmobiliare.xls"*/));
        this.fK = new HSSFWorkbook(fileKonta);
        this.sheetK = fK.getSheetAt(0);
        
        this.fileDanea =new  FileInputStream(new File(percorsoDanea/*"D:\\LavoroPapa\\FileDanea.xls"*/));
        this.fD = new HSSFWorkbook(fileDanea);
        this.sheetD = fD.getSheetAt(0);
        LetturaFileDanea.contaUtenti(sheetD);
        creaRighe();
    }
    
    public void inserisciDati() throws IOException{
        setScalaNumAlloggio();
        setNomeUtente();
        setCF();
        setNumTell();
        setEmail();
        setIndirizzo();
        setCivico();
        setScala();
        salvaFile();
    }
    private void creaRighe() {
         for(int i=6;i<LetturaFileDanea.contaUtenti(sheetD)*6;i++){
            sheetK.createRow(i);
        }   
    }
    private void setScalaNumAlloggio() throws IOException{
        int j=0;
        for(int i=6;i<LetturaFileDanea.contaUtenti(sheetD)*6;i=i+6){
            sheetK.getRow(i).createCell((short)0).setCellValue(LetturaFileDanea.getScalaNumAlloggio(sheetD).get(j));
            j++;
        }        
    }
    
    private void setNomeUtente() throws IOException{
        int j=0;
        for(int i=6;i<LetturaFileDanea.contaUtenti(sheetD)*6;i=i+6){
            sheetK.getRow(i).createCell((short)1).setCellValue(LetturaFileDanea.getNomeProprietario(sheetD).get(j));
            j++;
        }        
    }
    
    private void setCF() throws IOException{
        int j=0;
        for(int i=6;i<LetturaFileDanea.contaUtenti(sheetD)*6;i=i+6){
            sheetK.getRow(i).createCell((short)5).setCellValue(LetturaFileDanea.getCF(sheetD).get(j));
            j++;
        }        
    }
    
    private void setNumTell() throws IOException{
        int j=0;
        for(int i=6;i<LetturaFileDanea.contaUtenti(sheetD)*6;i=i+6){
            sheetK.getRow(i).createCell((short)2).setCellValue(LetturaFileDanea.getNumTell(sheetD).get(j));
            j++;
        }        
    }
    private void setEmail() throws IOException{
        int j=0;
        for(int i=6;i<LetturaFileDanea.contaUtenti(sheetD)*6;i=i+6){
            sheetK.getRow(i).createCell((short)6).setCellValue(LetturaFileDanea.getEmail(sheetD).get(j));
            j++;
        }        
    }
     private void setIndirizzo() throws IOException{
        int j=0;
        for(int i=6;i<LetturaFileDanea.contaUtenti(sheetD)*6;i=i+6){
            sheetK.getRow(i).createCell((short)7).setCellValue(LetturaFileDanea.getIndirizzo(sheetD).get(j));
            j++;
        }        
    }
     
    private void setCivico() throws IOException{
        int j=0;
        for(int i=6;i<LetturaFileDanea.contaUtenti(sheetD)*6;i=i+6){
            sheetK.getRow(i).createCell((short)8).setCellValue(LetturaFileDanea.getCivico(sheetD).get(j));
            j++;
        }        
    }
    
    private void setScala() throws IOException{
        int j=0;
        for(int i=6;i<LetturaFileDanea.contaUtenti(sheetD)*6;i=i+6){
            sheetK.getRow(i).createCell((short)9).setCellValue(LetturaFileDanea.getScala(sheetD).get(j));
            j++;
        }        
    }
    
    private void salvaFile(){
        try (FileOutputStream fos = new FileOutputStream(percorsoKonta/*"D:\\LavoroPapa\\ImportUnitàImmobiliare.xls"*/)) {
            fK.write(fos);
           // fK.close(); 
            fileDanea.close();
            //fD.close();
            System.out.println("File salvato correttamente.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
