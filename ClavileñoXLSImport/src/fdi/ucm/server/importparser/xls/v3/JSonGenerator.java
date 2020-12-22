package fdi.ucm.server.importparser.xls.v3;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.Properties;
import java.util.Scanner;

public class JSonGenerator {

	public static void main(String[] args) {
		
		String preLan="es";
		
		if (args.length>0)
			preLan=args[0].toLowerCase();
		
		Properties prop = new Properties();
		
		try (InputStream input = new FileInputStream("confxls.properties")) {
            prop.load(input);
	        } catch (IOException io) {
	            io.printStackTrace();
	        }
		
		LinkedList<File> ArchivosXLS=new LinkedList<File>();
		File carpeta = new File(".");
		File[] listado = carpeta.listFiles();
		if (listado == null || listado.length == 0) {
		    System.err.println(GetString(preLan,"no_elem_act",prop,"No hay elementos dentro de la carpeta actual"));
		    return;
		}

		    for (int i=0; i< listado.length; i++) {
//		        System.out.println(listado[i]);
		        if (listado[i].getName().endsWith(".xls")||listado[i].getName().endsWith(".xlsx"))
		        	ArchivosXLS.add(listado[i]);
		    }
		
		
		
		if (ArchivosXLS.size() == 0) {
			 System.err.println(GetString(preLan,"no_elem_xls_act",prop,"No hay elementos dentro de la carpeta actual"));
		    return;
		}
		
		
		
		
		
		System.out.println(GetString(preLan,"wellcome",prop,"Bienvenido al sistema generador de JSON para Importador XLS/XLSX Clavy"));
		System.out.println();
		
		boolean seleccion=false;
		
		Scanner myObj = new Scanner(System.in);
		
		File SelectedFile=null;
		
		while (!seleccion)
		{
		System.out.println(GetString(preLan,"seleccion",prop,"Por favor seleccione el archivo con el que desea trabajar"));
		
		HashMap<Integer, File> FileSelect=new HashMap<Integer, File>();
		int index=1;
		for (File fileXLS : ArchivosXLS) {
			FileSelect.put(new Integer(index), fileXLS);
			System.out.println(index+": "+fileXLS.getName());
			index++;
		}
		
		System.out.println(GetString(preLan,"input_abort",prop,"Introducir 0 para abortar"));
		System.out.println();
		
		Integer entrada=-1;
		
		try {
			entrada=Integer.parseInt(myObj.nextLine());
		} catch (Exception e) {
			// TODO: handle exception
		}
				
		if (entrada==0)
		{
			System.err.println(GetString(preLan,"input_abort_result",prop,"Operacion abortada"));
			return;
		}
		
		if (entrada>0&&entrada<index)
		{
			SelectedFile=FileSelect.get(new Integer(entrada));
			seleccion=true;
		}else
		{
			System.err.println(GetString(preLan,"input_error",prop,"Entrada no valida"));
			System.out.println();
		}
		
		
		
		}
		
		myObj.close();
		System.out.println(SelectedFile.getName());
	}

	private static String GetString(String preLan, String string, Properties prop,String Default) {
		if (prop.get(preLan+"_"+string)!=null)
			return prop.get(preLan+"_"+string).toString();
		else
			return Default;
	}
}
