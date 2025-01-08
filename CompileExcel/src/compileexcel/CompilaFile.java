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
import java.util.ArrayList;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
/**
 *
 * @author matti
 */
class CompilaFile {
    private FileInputStream fileDanea;
    private FileInputStream fileKonta;
    private HSSFWorkbook fK;
    private HSSFSheet sheetK;
    private String percorsoKonta;
    private String percorsoDanea;
    private HSSFWorkbook fD;
    private HSSFSheet  sheetD;
    
    public CompilaFile(String percorsoKonta,String percorsoDanea) throws FileNotFoundException, IOException{
        this.percorsoKonta=percorsoKonta;
        this.fileKonta = new  FileInputStream(new File(percorsoKonta/*"D:\\LavoroPapa\\ImportUnitàImmobiliare.xls"*/));
        this.fK = new HSSFWorkbook(fileKonta);
        this.sheetK = fK.getSheetAt(0);
        this.percorsoDanea=percorsoDanea;
        this.fileDanea =new  FileInputStream(new File(percorsoDanea/*"D:\\LavoroPapa\\FileDanea.xls"*/));
        this.fD = new HSSFWorkbook(fileDanea);
        this.sheetD = fD.getSheetAt(0);
        this.sheetK=GestoreExcel.creaRighe(sheetK,sheetD);
    }
    
    public ArrayList<DatiRiga> inserimentoInquilini() throws IOException{
        ArrayList<DatiRiga> righe=new ArrayList<>();
        for(int i=0;i<GestoreExcel.contaUtenti(sheetD, false);i++){
            righe.add(new DatiRiga(i,sheetD));       
        }
        
          for (int i = 0; i < righe.size(); i++) {
        DatiRiga riga = righe.get(i);
        if (!"Pr".equals(riga.ruolo)) {
            boolean trovato = false;
            for (int j = 0; j < righe.size(); j++) {
                if (i != j && riga.codScala.equals(righe.get(j).codScala) && "Pr".equals(righe.get(j).ruolo)) {
                    righe.get(j).setInquilino(riga.nome);
                    trovato = true;
                }
            }
            if (trovato) {
                righe.remove(i); // Rimuovi la riga non "Pr"
                i--; // Decrementa `i` per compensare la rimozione
            }
        }
    }
          return righe;
    }
    public void inserisciDati() throws IOException{
        for(int j=0;j<13;j++){
            for(int i=6;i<GestoreExcel.contaUtenti(sheetD,true)*6;i=i+6){
                switch(j){
                    case 0:
                        sheetK.getRow(i).createCell((short)j).setCellValue(/*new DatiRiga(i/6,sheetD).codScala*/inserimentoInquilini().get(i/6-1).codScala);
                        break;
                    case 3:
                        sheetK.getRow(i).createCell((short)j).setCellValue(inserimentoInquilini().get(i/6-1).nome);
                        break;
                    case 2:
                        sheetK.getRow(i).createCell((short)j).setCellValue(inserimentoInquilini().get(i/6-1).numTell);
                        break;
                    case 1:
                        sheetK.getRow(i).createCell((short)j).setCellValue(inserimentoInquilini().get(i/6-1).inquilino);
                        break;    
                    case 5:
                        sheetK.getRow(i).createCell((short)j).setCellValue(inserimentoInquilini().get(i/6-1).cf);
                        break;
                    case 6:
                        sheetK.getRow(i).createCell((short)j).setCellValue(inserimentoInquilini().get(i/6-1).email);
                        break;
                    case 7:
                        sheetK.getRow(i).createCell((short)j).setCellValue(inserimentoInquilini().get(i/6-1).indirizzo);
                        break;
                    case 8:
                        sheetK.getRow(i).createCell((short)j).setCellValue(inserimentoInquilini().get(i/6-1).civico);
                        break;
                    case 9:
                        sheetK.getRow(i).createCell((short)j).setCellValue(inserimentoInquilini().get(i/6-1).scala);
                        break;
                    case 13:
                        break;
                }
            }
        }
        salvaFile();
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
