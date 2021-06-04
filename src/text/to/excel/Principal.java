package text.to.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;
import java.time.LocalDate;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Principal {
	   public static void main(String args[]) throws Exception{

		      //Getting the Path object
		      String filePath = "C:\\Users\\YR9797\\Desktop\\prtugal_rejet_202105.txt";
		      Scanner sc = new Scanner(System.in);
		      System.out.println("********************************************************************************************************");
		      System.out.println("Entrez le path de fichier sous cette forme : C:\\Users\\YR9797\\Desktop\\prtugal_rejet_202105.txt");
		      filePath = sc.next();
		      File text = new File(filePath);
		      System.out.println("Lecture du fichier en cours ...");

		        //Creating Scanner instnace to read File in Java
		        Scanner scnr = new Scanner(new FileInputStream(text), "UTF-8");
		     
		        //Reading each line of file using Scanner class
		        int lineNumber = 1;
		        int info90 = 0;
		        int infodv = 0;
		        String matricule=null;
		        ArrayList <InformationDV> dvResult = new ArrayList<InformationDV>();
		        ArrayList <Information90> result90 = new ArrayList<Information90>();
		        while(scnr.hasNextLine()){
		        	try {
			            String line = scnr.nextLine();
			            String info = line.substring(3, 19);
			            String matcle = line.substring(3,48);
			            
			            if(matcle.contentEquals("STRUCTURE DE DONNEES :  ZY   CLE DU DOSSIER :")) {
			            	matricule = line.substring(53, 59);
			            }			          
			            if(info.contentEquals("INFORMATION : 90")) {
			            	info90++;
			            	Information90 inform90 = new Information90();
			            	scnr.nextLine();
			            	line = scnr.nextLine();
			            	inform90.setDossPaie(Integer.parseInt(line.substring(26,27)));
			            	inform90.setMatricule(matricule);
			            	inform90.setPERIODE((line.substring(67,69)+line.substring(110,114)).replaceAll(" ", ""));
			            	scnr.nextLine();
			            	line = scnr.nextLine();
			            	inform90.setRubrique(line.substring(67,70));
			            	inform90.setNUMORD(line.substring(112,125).replaceAll(" ", ""));
			            	result90.add(inform90);
			            }else if(info.contentEquals("INFORMATION : DV")) {
			            	scnr.nextLine();
			            	scnr.nextLine();
			            	line = scnr.nextLine();
			            	InformationDV infoDV = new InformationDV(matricule,line.substring(67, 77));
			            	dvResult.add(infoDV);
			            	infodv++;
			            }
			            lineNumber++;
		        	}catch(Exception ex) {
		        		
		        	}
		        } 
            	System.out.println("Lecture terminée");
            	System.out.println("*******************************************************************************************************");
  		      	System.out.println("Entrez le chemin du dossier destinataire sous cette forme : C:\\Users\\YR9797\\Desktop");
            	String filePath1 = "C:\\Users\\YR9797\\Desktop";
   		      	Scanner sc1 = new Scanner(System.in);
   		      	filePath1 = sc1.next();
            	System.out.println("Génération du fichier excel en cours ...");
            	
		         HSSFWorkbook workbook = new HSSFWorkbook();
		         // Create two sheets in the excel document and name it First Sheet and
		         HSSFSheet ZYDVSheet = workbook.createSheet("ZYDV");
		         HSSFSheet ZY90Sheet = workbook.createSheet("ZY90");
		         HSSFRow rowA = ZYDVSheet.createRow(0);
		         HSSFCell cellA = rowA.createCell(0);
		         HSSFCell cellB = rowA.createCell(1);
		         cellA.setCellValue(new HSSFRichTextString("Matricule"));
		         cellB.setCellValue(new HSSFRichTextString("Date fin consommation"));
		         for(int i=0;i<dvResult.size();i++) {
		        	   HSSFRow row = ZYDVSheet.createRow(i+1);
				         HSSFCell cell1 = row.createCell(0);
				         HSSFCell cell2 = row.createCell(1);
				         cell1.setCellValue(new HSSFRichTextString(dvResult.get(i).getMatricule()));
				         cell2.setCellValue(new HSSFRichTextString(dvResult.get(i).getDate()));
		         }

		         HSSFRow rowI = ZY90Sheet.createRow(0);
		         HSSFCell coll1 = rowI.createCell(0);
		         HSSFCell coll2 = rowI.createCell(1);
		         HSSFCell coll3 = rowI.createCell(2);
		         HSSFCell coll4 = rowI.createCell(3);
		         HSSFCell coll5 = rowI.createCell(4);
		         coll1.setCellValue(new HSSFRichTextString("Matricule"));
		         coll2.setCellValue(new HSSFRichTextString("Dossier de paie"));
		         coll3.setCellValue(new HSSFRichTextString("Rubrique"));
		         coll4.setCellValue(new HSSFRichTextString("PERIODE"));
		         coll5.setCellValue(new HSSFRichTextString("NUMORD"));
		         for(int i=0;i<result90.size();i++) {
			      	   HSSFRow row = ZY90Sheet.createRow(i+1);
				       HSSFCell cell1 = row.createCell(0);
				       HSSFCell cell2 = row.createCell(1);
				       HSSFCell cell3 = row.createCell(2);
				       HSSFCell cell4 = row.createCell(3);
				       HSSFCell cell5 = row.createCell(4);
				       cell1.setCellValue(new HSSFRichTextString(result90.get(i).getMatricule()));
				       String dossier = String.valueOf(result90.get(i).getDossPaie());
				       cell2.setCellValue(new HSSFRichTextString(dossier));
				       cell3.setCellValue(new HSSFRichTextString(result90.get(i).getRubrique()));
				       cell4.setCellValue(new HSSFRichTextString(result90.get(i).getPERIODE()));
				       cell5.setCellValue(new HSSFRichTextString(result90.get(i).getNUMORD()));
		         }
		         // To write out the workbook into a file we need to create an output
		         // stream where the workbook content will be written to.
		         String todaysDate = LocalDate.now().toString() .replaceAll(" ", "");
		         try (FileOutputStream fos = 
		                  new FileOutputStream(new File(filePath1+"\\PRT_REJETS"+todaysDate+".xls"))) {
		             workbook.write(fos);
		            	System.out.println("Fichier excel généré avec succès sur : "+filePath1+"\\PRT_REJETS-"+todaysDate+".xls");

		         } catch (IOException e) {
		             e.printStackTrace();
		         }
		         
		   }
}
