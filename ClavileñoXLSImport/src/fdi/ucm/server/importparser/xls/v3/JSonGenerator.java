package fdi.ucm.server.importparser.xls.v3;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map.Entry;
import java.util.Properties;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import fdi.ucm.server.importparser.xls.v3.CollectionXLS.FileFormat;
import fdi.ucm.server.importparser.xls.v3.struture.Hoja;
import fdi.ucm.server.importparser.xls.v3.struture.HojaV2;

public class JSonGenerator {

	enum FileFormat {OLD,NEW};
	
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
		
		
		
		
		 HashMap<String,HashMap<String,String>> Hojas=new HashMap<String,HashMap<String,String>>();

			
		 try {
			 if (SelectedFile.getName().endsWith(".xlsx")) 
				  Hojas=GENERAR(SelectedFile.getName(),FileFormat.NEW); 
			 else if (SelectedFile.getName().endsWith(".xls")) 
				  Hojas=GENERAR(SelectedFile.getName(),FileFormat.OLD);
		} catch (Exception e) {
			e.printStackTrace();
		}
		 
		 boolean continuarAdd=true;
		 
		 while (continuarAdd){
		  System.out.println();
		  boolean seleccion2=false;
		  String SelectedHoja=null;
		  while (!seleccion2)
			{
		 int index2=1;
		 HashMap<Integer, String> HojaSelect=new HashMap<Integer, String>();
		 System.out.println(GetString(preLan,"input_hoja",prop,"Seleccionar la hoja a procesar"));
		for (Entry<String, HashMap<String, String>> ee : Hojas.entrySet()) {
			HojaSelect.put(new Integer(index2), ee.getKey());
			System.out.println(index2+": "+ee.getKey());
			index2++;
		}

		Integer entrada=-1;
		
		try {
			entrada=Integer.parseInt(myObj.nextLine());
		} catch (Exception e) {
			// TODO: handle exception
		}
				
		
		if (entrada>0&&entrada<index2)
		{
			SelectedHoja=HojaSelect.get(new Integer(entrada));
			seleccion2=true;
		}else
		{
			System.err.println(GetString(preLan,"input_error",prop,"Entrada no valida"));
			System.out.println();
		}
		
			}
		  
		  String SelectedID=null;
		  if (SelectedHoja!=null)
		  {
			  boolean seleccion3=false;
			  
			  while (!seleccion3)
				{
			  System.out.println();
			  System.out.println(GetString(preLan,"input_columna_id",prop,"Seleccionar la columna con el identificador de CLAVY a procesar"));
			  HashMap<Integer, String> C_ID_Select=new HashMap<Integer, String>();
			  int index3=1;
			  for (Entry<String, String> colid : Hojas.get(SelectedHoja).entrySet()) {
				  C_ID_Select.put(new Integer(index3), colid.getKey());
					System.out.println(index3+": "+colid.getValue());
					index3++;
			  }
			  
			  Integer entrada=-1;
				
				try {
					entrada=Integer.parseInt(myObj.nextLine());
				} catch (Exception e) {
					// TODO: handle exception
				}
			  
			  
			  if (entrada>0&&entrada<index3)
				{
				  SelectedID=C_ID_Select.get(new Integer(entrada));
					seleccion3=true;
				}else
				{
					System.err.println(GetString(preLan,"input_error",prop,"Entrada no valida"));
					System.out.println();
				}
			  
				}
		  }
		  
		  String SelectedValor=null;
		  if (SelectedID!=null)
		  {
			  boolean seleccion4=false;
			  
			  while (!seleccion4)
				{
			  System.out.println();
			  System.out.println(GetString(preLan,"input_columna_id",prop,"Seleccionar la columna con el identificador de CLAVY a procesar"));
			  HashMap<Integer, String> C_ID_Select=new HashMap<Integer, String>();
			  int index3=1;
			  for (Entry<String, String> colid : Hojas.get(SelectedHoja).entrySet()) {
				  C_ID_Select.put(new Integer(index3), colid.getKey());
					System.out.println(index3+": "+colid.getValue());
					index3++;
			  }
			  
			  Integer entrada=-1;
				
				try {
					entrada=Integer.parseInt(myObj.nextLine());
				} catch (Exception e) {
					// TODO: handle exception
				}
			  
			  
			  if (entrada>0&&entrada<index3)
				{
				  SelectedValor=C_ID_Select.get(new Integer(entrada));
				  seleccion4=true;
				}else
				{
					System.err.println(GetString(preLan,"input_error",prop,"Entrada no valida"));
					System.out.println();
				}
			  
				}
		  }
		  
		  if (SelectedValor!=null)
		  {
			  System.out.println();
			  System.out.println(GetString(preLan,"confirm_values",prop,"Esta seguro de que desea añadir la siguiente informacion?"));
			  System.out.println(GetString(preLan,"pestana_confirm_values",prop,"Pestaña")+":"+SelectedHoja+ "  //  "+
					  GetString(preLan,"ide_confirm_values",prop,"Columna Identificador")+":"+Hojas.get(SelectedHoja).get(SelectedID)+ "  //  " +
					  GetString(preLan,"value_confirm_values",prop,"Columna Valor en caso de omision")+":"+Hojas.get(SelectedHoja).get(SelectedValor));
		  }
		  
		  
		  System.out.println();
		  System.out.println(GetString(preLan,"confirm_add_relacion",prop,"Desea añadir otra relacion? Introducir NO para parar"));
		  String continueStringIn=myObj.nextLine();
		  if (continueStringIn.toLowerCase().equals("no"))
			  continuarAdd=false;
  
		  if (continuarAdd)
			  System.out.println(GetString(preLan,"confirm_add_relacion",prop,"Desea añadir otra relacion? Introducir NO para parar"));
		  
	}
		  
		  myObj.close();
	}

	private static HashMap<String,HashMap<String,String>> GENERAR(String Nombre_Archivo, FileFormat FileFormatIn) throws IOException {

		HashMap<String,HashMap<String,String>> Salida=new HashMap<String,HashMap<String,String>>();
		 

			  Workbook Libro_trabajo;
			  
			  
			  /**
				 
			    * Crea una nueva instancia de la clase FileInputStream
			 
			    */
			 
			   FileInputStream fileInputStream = new FileInputStream(
			 
			     Nombre_Archivo);
			 
			  
			  if (FileFormatIn==FileFormat.NEW)
			  {

				   /**
				 
				    * Crea una nueva instancia de la clase XSSFWorkBook
				 
				    */
				 
				   Libro_trabajo = new XSSFWorkbook(fileInputStream);
			  }
			  else
			  {

				   POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
				   
				   Libro_trabajo = new HSSFWorkbook(fsFileSystem);
				 
			  }
		 
		 
			  int NStilos=Libro_trabajo.getNumberOfSheets(); 
			  
			  for (int i = 0; i < NStilos; i++) {
				  
				  Sheet Hoja_hssf = Libro_trabajo.getSheetAt(i);
				   
				   Iterator<Row> Iterador_de_Fila = Hoja_hssf.rowIterator();
					 
				   HashMap<String, String> SalidaPar = new HashMap<String,String>();
				   
				   Row Fila_hssf = (Row) Iterador_de_Fila.next();
				   

					   Iterator<Cell> iterador = Fila_hssf.cellIterator();
				    	while (iterador.hasNext()) {
				    		Cell hssfCell = (Cell) iterador.next();
				    		String column_letter = CellReference.convertNumToColString(hssfCell.getColumnIndex());
						    
				    		 String Valor_de_celda = "~uname";
							 

						     
						     if(hssfCell.getCellType() == Cell.CELL_TYPE_FORMULA){
						    	 switch(hssfCell.getCachedFormulaResultType()) {
						            case Cell.CELL_TYPE_NUMERIC:
						                System.out.println("Last evaluated as: " + hssfCell.getNumericCellValue());
						                Valor_de_celda=Double.toString(hssfCell.getNumericCellValue());
						                break;
						            case Cell.CELL_TYPE_STRING:
						                System.out.println("Last evaluated as \"" + hssfCell.getRichStringCellValue() + "\"");
						                Valor_de_celda=hssfCell.getRichStringCellValue().toString();
						                break;
						             default:
						            	Valor_de_celda = hssfCell.toString();
								        break;
						                	
						        }
						     }else
						    	 Valor_de_celda = hssfCell.toString();
						 
						 
						     SalidaPar.put(column_letter, Valor_de_celda);	
				    	
				    	}
				    	
				   
				 Salida.put(Hoja_hssf.getSheetName(), SalidaPar); 
			  }
		
		return Salida;
	}

	private static String GetString(String preLan, String string, Properties prop,String Default) {
		if (prop.get(preLan+"_"+string)!=null)
			return prop.get(preLan+"_"+string).toString();
		else
			return Default;
	}
}
