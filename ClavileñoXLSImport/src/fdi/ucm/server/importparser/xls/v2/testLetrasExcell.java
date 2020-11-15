package fdi.ucm.server.importparser.xls.v2;

import java.util.Scanner;

public class testLetrasExcell {

	public static void main(String[] args) {
		Scanner myObj = new Scanner(System.in);  // Create a Scanner object

	    String userName = myObj.nextLine();  // Read user input
	    System.out.println("Username is: " + userName);  // Output user input
	    
	    userName=userName.toUpperCase();
	    
	    Integer Final=0;
	    
	    int BaseZ = 26;
	    
	    for (int i=0; i<userName.length();i++)
	    {
	        int pos = userName.length()-1-i;
	        char charar = userName.charAt(pos);
	        
	        Integer I = ((int)(new Character(charar))) - ((int)(new Character('A')))+1;
	        
	        
	       double Var = Math.pow(BaseZ, pos)*I;

	        Final=Final+(int)Var;
	    }
	    
	    System.out.println("->>>"+(Final-1));
	    myObj.close();
	}

}
