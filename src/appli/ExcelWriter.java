package appli;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;


import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;


import static java.lang.System.out;

public class ExcelWriter {


    private String txt;
    private HSSFWorkbook wb;
    private List<List<String>> sheetL;
    private ExcelReader excelR;

    ExcelWriter(String txt, HSSFWorkbook wb, List<List<String>> sheetL,ExcelReader excelR) {
        this.txt = txt;
        this.wb = wb;
        this.sheetL = sheetL;
        this.excelR=excelR;
    }

    void generateExcel(){
        DecimalFormat df = new DecimalFormat("000.00");

        HSSFSheet sheet = wb.createSheet(this.txt);
        List<List<String>> newSheetL = new ArrayList<>();
        int cpt=0;
        List<String> newRowL ;
        float oldInt;
        float oldInt8;
        float oldInt10;
        excelR.getBar().setValue(30);
        for (List<String> oldRow : sheetL ){
            if (cpt == 0) {

                List<String> newRowL1 = new ArrayList<>();
                newRowL1.add("");
                newRowL1.add("POIDS CLIENT");
                newRowL1.add("POIDS CLIENT");
                newRowL1.add("ALL Evolution");
                newRowL1.add("ALL Evolution");
                newRowL1.add("Ebiz Evolution");
                newRowL1.add("Ebiz Evolution");
                newRowL1.add("B2B Evolution");
                newRowL1.add("B2B Evolution");
                newRowL1.add("WEB Evolution");
                newRowL1.add("WEB Evolution");
                newRowL1.add("Monitoring");


                List<String> newRowL2 = new ArrayList<>();
                newRowL2.add("Composite Code ");
                newRowL2.add("Lignes n-1");
                newRowL2.add("Lignes n");
                newRowL2.add("%");
                newRowL2.add("Lignes");
                newRowL2.add("%");
                newRowL2.add("%Globale");
                newRowL2.add("%");
                newRowL2.add("%Globale");
                newRowL2.add("%");
                newRowL2.add("%Globale");
                newRowL2.add("Monitoring");

                newSheetL.add(newRowL1);
                newSheetL.add(newRowL2);


                cpt++;
                excelR.getBar().setValue(35);
                excelR.repaint();
            }

            else if(cpt == 1){
                newRowL = new ArrayList<>();
                newSheetL.add(newRowL);
                cpt++;
            }

            //après les deux lignes d'entete
            else{//ALGORITHM DE CREATION DES CASES
                newRowL = new ArrayList<>();
                float b=0;
                float c=0;
                float all=0;
                float ebiz=0;

                Boolean isNew = false;
                for (int i = 0; i < 12; i++) {
                    switch (i) {
                        case 0:
                            //Composite Code
                            newRowL.add(oldRow.get(0));
                            System.out.println(oldRow.get(0));
                            break;
                        case 1:
                            //ligne n-1
                            b = Float.parseFloat(oldRow.get(5));

                            newRowL.add(oldRow.get(5));
                            break;
                        case 2:
                            //ligne n
                            c = Float.parseFloat(oldRow.get(1));

                            newRowL.add(oldRow.get(1));
                            break;
                        case 3:
                            //%
                            float value3 = ((c - b) / b)*100;
                            if (b == 0) {
                                newRowL.add("NEW");
                                isNew = true;
                            }
                            else if(c == 0){
                                newRowL.add("-001,0");
                            }
                            else {
                                newRowL.add(df.format((value3)));
                                all = value3;
                            }
                            break;
                        case 4:
                            //Lignes

                            float leC=Float.parseFloat(oldRow.get(1));
                            System.out.println("C : "+leC);
                            float leB=Float.parseFloat(oldRow.get(5));
                            System.out.println("b : "+leB);
                            float leE=leC-leB;
                            System.out.println("e : "+leE);
                            String newCS=Float.toString(leE);
                            newRowL.add(newCS);
                            break;
                        //ebiz
                        case 5:
                            //%
                            float oldC;
                            float oldG;

                            String oldCS = oldRow.get(2);
                            System.out.println("///////////////////OLDCS : "+oldCS);
                            if (oldCS.equals("-")){
                                oldC=0;
                                System.out.println(" - : 0" );
                            }
                            else {
                                oldC = Float.parseFloat(oldCS);
                                System.out.println("else : "+oldC);
                            }

                            String oldGS = oldRow.get(6);
                            System.out.println("/////////////////oldGS : "+oldGS);
                            if (oldGS.equals("-")){
                                oldG=0;
                                System.out.println(" - : 0" );
                            }
                            else {
                                System.out.println(oldGS.indexOf('E')+"//////////////////////////////////////");
                                if (oldGS.indexOf('E') != -1){
//                                    oldG = Float.parseFloat(oldGS.split("E")[0])/100;
                                    oldG = Float.valueOf(oldGS);
                                }
                                else {
                                    oldG=Float.parseFloat(oldGS);
                                }

                            }
                            float value5 = (oldC - oldG)*100;
                            ebiz = value5;

                            newRowL.add(df.format((value5) ));
                            break;
                        case 6:
                            //global
                            String old = oldRow.get(2);

                            if (old.equals("-")){
                                oldInt=0;
                            }
                            else{

                                if (old.indexOf('E') != -1){
//                                    oldInt = Float.parseFloat(old.split("E")[0])/100;
                                    oldInt = Float.valueOf(old);
                                }
                                else {
                                    oldInt=Float.parseFloat(old);
                                }


                            }

                            newRowL.add(Float.toString(oldInt*100));
                            break;
                        //B2B
                        case 7:
                            //%
                            float oldD7;
                            float oldH7;

                            String oldDS7 = oldRow.get(3);
                            if (oldDS7.equals("-")){
                                oldD7=0;
                            }
                            else {

                                if (oldDS7.indexOf('E') != -1){
//                                    oldD7 = Float.parseFloat(oldDS7.split("E")[0])/100;
                                    oldD7 = Float.valueOf(oldDS7);
                                }
                                else {
                                    oldD7 = Float.parseFloat(oldDS7);
                                }
                            }

                            String oldHS7 = oldRow.get(7);
                            if (oldHS7.equals("-")){
                                oldH7=0;
                            }
                            else {

                                if (oldHS7.indexOf('E') != -1){
//                                    oldH7 = Float.parseFloat(oldHS7.split("E")[0])/100;
                                    oldH7 = Float.valueOf(oldHS7);
                                }
                                else {
                                    oldH7 = Float.parseFloat(oldHS7);
                                }
                            }
                            float value7 = (oldD7 - oldH7)*100;

                            newRowL.add(df.format((value7)));
                            break;
                        case 8:
                            //global

                            String oldD8 = oldRow.get(3);
                            if (oldD8.equals("-")){
                                oldInt8=0;
                            }
                            else{

                                if (oldD8.indexOf('E') != -1){
//                                    oldInt8 = Float.parseFloat(oldD8.split("E")[0])/100;
                                    oldInt8 = Float.valueOf(oldD8);
                                }
                                else {
                                    oldInt8 = Float.parseFloat(oldD8);
                                }
                            }

                            newRowL.add(Float.toString(oldInt8*100));

                            break;
                        //WEB
                        case 9:
                            //%
                            float oldE9;
                            float oldI9;

                            String oldES9 = oldRow.get(4);
                            if (oldES9.equals("-")){
                                oldE9=0;
                            }
                            else {

                                if (oldES9.indexOf('E') != -1){
//                                    oldE9 = Float.parseFloat(oldES9.split("E")[0])/100;
                                    oldE9 = Float.valueOf(oldES9);
                                }
                                else {
                                    oldE9 = Float.parseFloat(oldES9);
                                }
                            }

                            String oldIS9 = oldRow.get(8);
                            if (oldIS9.equals("-")){
                                oldI9=0;
                            }
                            else {

                                if (oldIS9.indexOf('E') != -1){
//                                    oldI9 = Float.parseFloat(oldIS9.split("E")[0])/100;
                                    oldI9 = Float.valueOf(oldIS9);
                                }
                                else {
                                    oldI9 = Float.parseFloat(oldIS9);
                                }
                            }
                            float value9 = (oldE9 - oldI9)*100;
                            System.out.println("E ("+oldE9+") - I ("+oldI9+") = "+value9);

                            newRowL.add(df.format((value9)));
                            break;
                        case 10:
                            //global

                            String old10 = oldRow.get(4);
                            if (old10.equals("-")){
                                oldInt10=0;
                            }
                            else{

                                if (oldRow.get(3).indexOf('E') != -1){
                                    oldInt10 = Float.parseFloat(oldRow.get(4).split("E")[0])/100;
                                }
                                else {
                                    oldInt10 = Float.parseFloat(oldRow.get(4));
                                }

                            }

                            newRowL.add(Float.toString(oldInt10*100));

                            break;
                        case 11:
                            String code;
                            //monitorCode
                            if (!isNew) {
                                if (all > 0) {
                                    if (ebiz > 0) {
                                        code = "1";
                                    } else if (ebiz == 0) {
                                        code = "2";
                                    } else {
                                        code = "3";
                                    }
                                } else if (all == 0) {
                                    if (ebiz > 0) {
                                        code = "4";
                                    } else if (ebiz == 0) {
                                        code = "5";
                                    } else {
                                        code = "6";
                                    }
                                } else {
                                    if (ebiz > 0) {
                                        code = "7";
                                    } else if (ebiz == 0) {
                                        code = "8";
                                    } else {
                                        if (all < ebiz) {
                                            code = "9";
                                        } else if (all == ebiz) {
                                            code = "10";
                                        } else {
                                            code = "11";
                                        }
                                    }
                                }
                            } else {
                                code = "0";
                            }
                            isNew = false;

                            newRowL.add(code);
                            break;
                    }

                }newSheetL.add(newRowL);
            }

        }
//        excelR.getBar().setValue(50);
        out.println("---------------------------sheet---------------------");

        int rowCount = 0;

        //Style du header
        HSSFFont font= wb.createFont();
        CellStyle style = wb.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        (style).setFillBackgroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        font.setColor(IndexedColors.BLACK.getIndex());
        font.setBold(true);
        style.setFont(font);

//        //Style pourcentage
        CellStyle stylePourcent = wb.createCellStyle();
//        stylePourcent.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
        stylePourcent.setDataFormat(wb.createDataFormat().getFormat("##0.0#"));



        for (List<String> unRow: newSheetL) {
//            out.println("1st");
            Row row = sheet.createRow(rowCount);
            rowCount++;
            int columnCount = 0;
            for (String unString : unRow) {
//                out.println("2nd");
                Cell cell = row.createCell(columnCount);
                if(rowCount<=3){
                    cell.setCellStyle(style);
                }
                if(rowCount>3 && (columnCount==3 ||columnCount==5||columnCount==6||columnCount==7||columnCount==8||columnCount==9||columnCount==10)){
                    cell.setCellStyle(stylePourcent);
                }
                columnCount++;
                if (isFloatable(unString)){
                    cell.setCellValue( java.lang.Float.parseFloat(unString.replace(",",".")));
                }
                else{
                    cell.setCellValue(unString);
                }
            }
        }
        excelR.getBar().setValue(80);
        sheet.setColumnWidth(0,9000);
        for (int g = 1;g<12;g++){
            sheet.setColumnWidth(g,3000);
        }

        try {
            FileOutputStream fos = new FileOutputStream(System.getProperty("user.home") + "/Desktop/xslCree.xls");
            wb.write(fos);
            fos.close();
            out.println("enregistré");
            excelR.addToConsole("Fichier enregistré");
        } catch (IOException e) {
//            e.printStackTrace();
//            this.excelR.addToConsole(e.toString());
        }
        excelR.getBar().setValue(90);
    }

    //GETTERS SETTERS
    public String getTxt() {
        return txt;
    }

    public void setTxt(String txt) {
        this.txt = txt;
    }

    public HSSFWorkbook getWb() {
        return wb;
    }

    public void setWb(HSSFWorkbook wb) {
        this.wb = wb;
    }

    public List getSheetL() {
        return sheetL;
    }

    private Boolean isFloatable(String value){
        value=value.replace('.', ',');
        Float valueF;
        if (value==""){
            return false;
        }
        try{
            valueF = Float.parseFloat(value.replace(",","."));
            return true;
        }catch (NumberFormatException e){
//            System.out.println(e);
//            excelR.addToConsole(e.toString());
            return false;
        }


    }

}
