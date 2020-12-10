package distance;
import java.util.Scanner;
import  java.io.*;
import  org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Main {
	
	public static void main(String[] args) throws IOException {
		
		//create the excel file "PlayerData" on desktop
		File fileName = new File("C:\\Users\\hadas\\Desktop\\ PlayerData.xls");
		
		//check if file exists already. If not, create the file
		if(!fileName.exists()) {
			fileName.createNewFile();
			
			//initializing output stream for writing data to the file
			FileOutputStream initializingStream = new FileOutputStream(fileName);
			
			//create object of a workbook and write to the file
			HSSFWorkbook initialWorkbook = new HSSFWorkbook();		
			initialWorkbook.write(initializingStream);
			
			//close stream and workbook
			initializingStream.close();
			initialWorkbook.close();	
			
			System.out.println("Your excel file has been generated!");
		}
		
		//control variable for input
		boolean checkInput = false;
		while (checkInput == false ) {
		System.out.println("Enter \"Add\" to add new player, \"Edit\" to edit existing player, or \"Delete\" to delete player's data");
		System.out.println("Enter \"Exit\" to terminate program");
		
		//take input for action
		Scanner input = new Scanner(System.in);
		String decision = input.next();
	
			//add info for existing player (each player has a sheet)
			if (decision.compareTo("edit") == 0 || decision.compareTo("Edit") == 0) {
				Editing.edit(fileName);		
			}
			//add new player (sheet)
			else if (decision.compareTo("add") == 0 || decision.compareTo("Add") == 0) {
				Player.addPlayer(fileName);
				Editing.edit(fileName);	
			}
			else if (decision.compareTo("delete") == 0 || decision.compareTo("Delete") == 0) {
				Editing.delete(fileName);	
			}
			else if (decision.compareTo("Exit") == 0 || decision.compareTo("exit") == 0) {
				System.exit(0);
				checkInput = true;
			}
			else {
				System.out.println("Invalid input, please try again");
				checkInput = false;
			}
		}
	}
}
