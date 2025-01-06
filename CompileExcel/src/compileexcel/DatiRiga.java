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
public class DatiRiga extends LetturaFileDanea{
    protected String codScala;
    protected String nome;
    protected String cf;
    protected String email;
    protected String indirizzo;
    protected String civico;
    protected String scala;
    protected String numTell;
    protected String ruolo;
    
    public DatiRiga(int nRiga,HSSFSheet  sheetD) throws IOException{
        super(sheetD);
        creaRiga(nRiga);
    }
    
    public void creaRiga(int nRiga) throws IOException{
        codScala=getScalaNumAlloggio().get(nRiga);
        nome=getNome().get(nRiga);
        cf=getCF().get(nRiga);
        email=getEmail().get(nRiga);
        indirizzo=getIndirizzo().get(nRiga);
        civico=getCivico().get(nRiga);
        scala=getScala().get(nRiga);
        numTell=getNumTell().get(nRiga);
        ruolo=getRuolo().get(nRiga);
    }
}
