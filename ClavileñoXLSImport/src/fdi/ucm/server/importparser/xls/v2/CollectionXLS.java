/**
 * 
 */
package fdi.ucm.server.importparser.xls.v2;

import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

import fdi.ucm.server.importparser.xls.v2.struture.Hoja;
import fdi.ucm.server.importparser.xls.v2.struture.HojaV2;
import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteTextElement;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteTextElementType;

import java.io.FileInputStream;
import java.io.FileReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;


/**
 * Clase que implementa la creacion de la base de datos per se
 * @author Joaquin Gayoso-Cabada
 *
 */
public class CollectionXLS implements InterfaceXLSparser {


	enum FileFormat {OLD,NEW};
	
	private static final String XLS_COLLECTION = "XLS COllection";
	private static final String COLECCION_A_APARTIR_DE_UN_XLS = "Coleccion a partir de un XLS";
	private CompleteCollection coleccionstatica;
	
	public CollectionXLS() {
		coleccionstatica=new CompleteCollection(XLS_COLLECTION, COLECCION_A_APARTIR_DE_UN_XLS+ new Timestamp(new Date().getTime()));
	}
	
	/* (non-Javadoc)
	 * @see fdi.ucm.server.importparser.sql.SQLparserModel#ProcessAttributes()
	 */
	@Override
	public void ProcessAttributes() {
		
	}

	
	 /**
	 
	  * Este metodo es usado para leer archivos Excel
	 
	  *
	 
	  * @param Nombre_Archivo
	 
	  *            - Nombre de archivo Excel.
	 * @param fileNameP 
	 
	  */
	 public void Leer_Archivo_Excel(String Nombre_Archivo, String fileNameP) {
	 
	  /**
	 
	   * Crea una nueva instancia de Lista_Datos_Celda
	 
	   */
		 
		 HashMap<String, List<Integer>> LinkElements=new HashMap<String, List<Integer>>();
		 
		 try {

			 JSONParser parser = new JSONParser();

			 Object obj = parser.parse(new FileReader(fileNameP));
					 
			 JSONArray jsonObject = (JSONArray) obj;
			 
			 for (int i = 0; i < jsonObject.size(); i++) {
				JSONObject array_element = (JSONObject) jsonObject.get(i);
				String Gram=array_element.get("g").toString();
				String COL_S=array_element.get("c").toString();
				
				if (Gram!=null&&COL_S!=null&&Gram.trim().length()>0&&COL_S.trim().length()>0)
					{
					Character COL_C=COL_S.toUpperCase().charAt(0);
					Integer I = ((int)(COL_C)) - ((int)(new Character('A')));
					
					List<Integer> listaPre= LinkElements.get(Gram);
					if (listaPre==null)
						listaPre=new LinkedList<Integer>();
					
					listaPre.add(I);
					LinkElements.put(Gram, listaPre);
					
					}
				
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	 
		 
	
		 
		 
	  ArrayList<Hoja> Hojas=new ArrayList<Hoja>();

	  
	 
	  if (Nombre_Archivo.contains(".xlsx")) 
		  Hojas=GENERAR(Nombre_Archivo,FileFormat.NEW); 
	 else if (Nombre_Archivo.contains(".xls")) 
		  Hojas=GENERAR(Nombre_Archivo,FileFormat.OLD);

	 
	  Imprimir_Consola(Hojas,LinkElements);
	 
	 }
	 
	 private ArrayList<Hoja> GENERAR(String Nombre_Archivo, FileFormat FileFormatIn) {
		 
		 ArrayList<Hoja> Salida=new ArrayList<Hoja>();
		 
	  try {
	 
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
	   int columProcess=0;
		 
	   for (int i = 0; i < NStilos; i++) {
		   
		   Sheet Hoja_hssf = Libro_trabajo.getSheetAt(i);
		   
		   HojaV2 Hojax=new HojaV2(Hoja_hssf.getSheetName());
		   
		   Iterator<Row> Iterador_de_Fila = Hoja_hssf.rowIterator();
			 
		   List<List<Cell>> Lista_Datos_Celda2 = new ArrayList<List<Cell>>();
		   
		   boolean primera=true;
		   while (Iterador_de_Fila.hasNext()) {
		 
			   
			   
			   Row Fila_hssf = (Row) Iterador_de_Fila.next();
			   
			   
			   List<Cell> Lista=new LinkedList<>();
			   
		 
			   if (primera)
			    {
				   Iterator<Cell> iterador = Fila_hssf.cellIterator();
			    	while (iterador.hasNext()) {
			    		Cell Celda_hssf = (Cell) iterador.next();
			    		Lista.add(Celda_hssf);
					 
					    }
			    	
			    	columProcess=Lista.size();
			    	
			    	primera=false;
			    }
			    else
			    {
			    	for (int j = 0; j < columProcess; j++) {
			    		Cell Celda_hssf=Fila_hssf.getCell(j);
			    		Lista.add(Celda_hssf);
					}
			    }
			   

		 
		    Lista_Datos_Celda2.add(Lista);
		 
		   }
		   
		   Hojax.setListaHijos(Lista_Datos_Celda2);
		   Salida.add(Hojax);
	}
	   
	   

	  } catch (Exception e) {
	 
	   e.printStackTrace();
	 
	  }
	 
	  return Salida;
	 }
	 
	 
	 
	 
	 /**
	 
	  * Este método se utiliza para imprimir los datos de la celda a la consola.
	 * @param linkElements 
	 
	  *
	 
	  * @param Datos_celdas
	 
	  *            - Listado de los datos que hay en la hoja de cálculo.
	 
	  */
	 
	 private void Imprimir_Consola(List<Hoja> HojasEntrada, HashMap<String, List<Integer>> linkElements) {
	 
		 
	
	for (Hoja hoja : HojasEntrada) {
		
//		System.out.println("Nombre: " + hoja.getName());
		
		CompleteGrammar Grammar=new CompleteGrammar(hoja.getName(), hoja.getName(), coleccionstatica);
		coleccionstatica.getMetamodelGrammar().add(Grammar);
		HashMap<Integer, CompleteTextElementType> Hash=new HashMap<Integer, CompleteTextElementType>();
		HashMap<String, CompleteTextElementType> HashPath=new HashMap<String, CompleteTextElementType>();
		
		 CompleteTextElementType Descripccion=null;
		 CompleteTextElementType Icon=null;
		
		 List<List<Cell>> Datos_celdas = ((HojaV2) hoja).getListaHijos();
			
		 String Valor_de_celda;
		 
		
		 
		  for (int i = 0; i < Datos_celdas.size(); i++) {
		 
			CompleteDocuments Doc=new CompleteDocuments(coleccionstatica, Integer.toString(i), "");  
			if (i!=0)
				coleccionstatica.getEstructuras().add(Doc);
			  
		   List<Cell> Lista_celda_temporal = Datos_celdas.get(i);
		 
		   for (int j = 0; j < Lista_celda_temporal.size(); j++) {
		 
		 
		     Cell hssfCell = Lista_celda_temporal.get(j);
		 
		     
		     Valor_de_celda="";
			 
			   if (Lista_celda_temporal.get(j)!=null)
			   {
		     
		     
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
		 
		 
		 
		    if (i==0)
		    	 {
		    	if (Valor_de_celda==null||Valor_de_celda.isEmpty())
		    		Valor_de_celda=hoja.getName()+" Columna:"+j;
		    	
		    	if (linkElements.get(hoja.getName()) != null&&linkElements.get(hoja.getName()).contains(new Integer(j)))
		    		System.out.println(j+":"+Valor_de_celda+"->> Deberia ser un Link");
		    	
		    	CompleteTextElementType C=generaStructura(Valor_de_celda,Grammar,HashPath);
		    	Hash.put(new Integer(j), C);
//		    	System.out.print("Columna:" + Valor_de_celda + "\t\t");
		    	
		    	if (isDesscription(Valor_de_celda))
		    		Descripccion=C;
		    	
		    	if (isIcon(Valor_de_celda))
		    		Icon=C;

		    	 }
		    
		    else 
		    	{
		    	CompleteTextElementType C=Hash.get(new Integer(j));
		    	if (C==null)
	    		{
		    	String Valor_de_celdaT = hoja.getName()+" Columna:"+j;
		    	C=generaStructura(Valor_de_celdaT,Grammar,HashPath);
		    	Hash.put(new Integer(j), C);
	    		}
		    	CompleteTextElement CT=new CompleteTextElement(C, Valor_de_celda);
		    	Doc.getDescription().add(CT);
//		    	System.out.print("Valor:" + Valor_de_celda + "\t\t");
		    	
		    	if (C==Descripccion)
		    		Doc.setDescriptionText(Valor_de_celda);
		    	
		    	if (C==Icon)
		    		Doc.setIcon(Valor_de_celda);
		    	
		    	}
			   }
		   }
		 
//		   System.out.println();
		 
		  }
		 
		
		
		
	}	 
		 
		 
	 
	 
	 }
	 
	 private boolean isIcon(String valor_de_celda) {
		 String comprara=valor_de_celda.trim().toLowerCase();

		 List<String> IconText=new LinkedList<>();
		 
		 IconText.add("icon");
		 IconText.add("ico");
		
		 
		if (IconText.contains(comprara))
			return true;
		else
			return false;
	}

	private boolean isDesscription(String valor_de_celda) {
		 String comprara=valor_de_celda.trim().toLowerCase();

		 List<String> DescrpitionText=new LinkedList<>();
		 
		 DescrpitionText.add("description");
		 DescrpitionText.add("desc");
		
		 
		if (DescrpitionText.contains(comprara))
			return true;
		else
			return false;
	}

	private CompleteTextElementType generaStructura(String valor_de_celda, CompleteGrammar grammar, HashMap<String, CompleteTextElementType> hashPath) {
		 
		
		 CompleteTextElementType preproducido = hashPath.get(valor_de_celda);
			if (preproducido!=null)
				return preproducido;
		 
		 
		String[] pathL=valor_de_celda.split("/");
		
		CompleteElementType Padre=null;
		
		 if (pathL.length>1)
			 Padre=producePadre(pathL,hashPath,grammar);
		 
		 CompleteTextElementType Salida=null;
		if (Padre!=null)
		 {
			Salida=new CompleteTextElementType(pathL[pathL.length-1], Padre, grammar);
			Padre.getSons().add(Salida);
		 }
		else 
			{
			Salida=new CompleteTextElementType(valor_de_celda, grammar);
			grammar.getSons().add(Salida);
			}
		
		hashPath.put(valor_de_celda, Salida);
		return Salida;
	}

	private CompleteElementType producePadre(String[] pathL,
			HashMap<String, CompleteTextElementType> hashPath,CompleteGrammar CG) {
		
		String Acumulado = "";
		CompleteTextElementType Padre = null;
		for (int i = 0; i < pathL.length-1; i++) {
			if (i!=0)
				Acumulado=Acumulado+"/"+pathL[i];
			else
				Acumulado=Acumulado+pathL[i];
			CompleteTextElementType yo = hashPath.get(Acumulado);
			if (yo==null)
				{
				
				if (Padre!=null)
					{
					CompleteTextElementType Salida = new CompleteTextElementType(pathL[i], Padre , CG);
					Padre.getSons().add(Salida);
					hashPath.put(Acumulado, Salida);
					}
				else
					{
					CompleteTextElementType Salida = new CompleteTextElementType(pathL[i], CG);
					CG.getSons().add(Salida);
					hashPath.put(Acumulado, Salida);
					}
				
				}
			
			Padre=yo;
		}
		return Padre;
	}

	public static void main(String[] args) {
	 
		 String fileName = "test.xls";
		if (args.length>0)
			fileName=args[0];
		
		String fileNameP = "test.xls.json";
		if (args.length>1)
			fileNameP=args[1];
	 
	  System.out.println(fileName);
	  
	  System.out.println(fileNameP);
	 
	 CollectionXLS C = new CollectionXLS();
	 C.Leer_Archivo_Excel(fileName,fileNameP);
	 
	 System.out.println(C.toString());
	 }

	@Override
	public void ProcessInstances() {
		
		
	}


	public CompleteCollection getColeccion() {
		return coleccionstatica;
	}
	
	
}
