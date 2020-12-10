package distance;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;

public class Player {

	//method to check if player sheet already exists
	public static String playerExist(HSSFWorkbook book, Scanner input) {
		System.out.println("Enter player's name: ");
		String playerName = input.nextLine();
		
		if(book.getNumberOfSheets() == 0) {
			return playerName;
		}
		else {
			boolean checkExist = false;
			while(checkExist == false) {
				for(int i = 0; i < book.getNumberOfSheets(); i++) {
					for(Sheet individualSheet: book) {
						String sheet = individualSheet.getSheetName();
						if(sheet.compareTo(playerName) == 0) {
							System.out.println("Player's name already exist. Try again");
							playerName = input.nextLine();
						}
						else {
							checkExist =  true;
						}
					}
				}
			}			
		}
		return playerName;
	}

	//method to add player sheet called when user wants to create new player 
	public static void addPlayer(File inputFile) throws IOException {
			 
			   //input stream object to read from the file
	    	   FileInputStream fileIn = new FileInputStream(inputFile);
	           
	    	   //new workbook for the file
	            HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
	            
	            //assigning sheet to player
	            Scanner input = new Scanner(System.in);
				String playerName = playerExist(workbook,input);
				
	            //create a sheet in the workbook set the sheet's name to player's name
	            HSSFSheet sheet = workbook.createSheet(playerName);
	            
	            //create the first row in the current sheet
	            HSSFRow rowhead = sheet.createRow((short)0);
	           
	            //set up color, font and alignment of the first row
	            HSSFCellStyle style = workbook.createCellStyle();
	            HSSFFont font = workbook.createFont();
	            font.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
	            font.setBold(true);
	            style.setFont(font);
	            style.setAlignment(HorizontalAlignment.CENTER);
	            
	            //creating and setting a header for each iron 
	            rowhead.createCell(0).setCellValue("4 Iron");
	            rowhead.createCell(1).setCellValue("5 Iron");
	            rowhead.createCell(2).setCellValue("6 Iron");
	            rowhead.createCell(3).setCellValue("7 Iron");
	            rowhead.createCell(4).setCellValue("8 Iron");
	            rowhead.createCell(5).setCellValue("9 Iron");
	            rowhead.createCell(6).setCellValue("P");
	            
	            //assign color, font and alignment for each header
	            for(int i = 0; i<=6; i++)
	            	rowhead.getCell(i).setCellStyle(style);

	           //initializing output stream for writing data to the file
	           FileOutputStream fileOut = new FileOutputStream(inputFile);

	           //write to sheet
	           workbook.write(fileOut);
	           //close streams and workbook
	           fileOut.close();
	           workbook.close();
	           fileIn.close();
	           
	           System.out.println("Your player sheet has been generated!");
	}
}

