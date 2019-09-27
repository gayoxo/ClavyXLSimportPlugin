/**
 * 
 */
package fdi.ucm.server.importparser.xls;

import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

import fdi.ucm.server.importparser.xls.struture.Hoja;
import fdi.ucm.server.importparser.xls.struture.HojaAntigua;
import fdi.ucm.server.importparser.xls.struture.HojaNueva;
import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteTextElement;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteTextElementType;

import java.io.FileInputStream;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Clase que implementa la creacion de la base de datos per se
 * @author Joaquin Gayoso-Cabada
 *
 */
public class CollectionXLS implements InterfaceXLSparser {


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
	 
	  */
	 public void Leer_Archivo_Excel(String Nombre_Archivo) {
	 
	  /**
	 
	   * Crea una nueva instancia de Lista_Datos_Celda
	 
	   */
	 
	  ArrayList<Hoja> Hojas=new ArrayList<Hoja>();

	 
	  if (Nombre_Archivo.contains(".xlsx")) {
	 
		  Hojas=GENERAR_XLSX(Nombre_Archivo);
	 
	  } else if (Nombre_Archivo.contains(".xls")) {
	 
		  Hojas=GENERAR_XLS(Nombre_Archivo);

	 
	  }
	 
	  Imprimir_Consola(Hojas);
	 
	 }
	 
	 private ArrayList<Hoja> GENERAR_XLSX(String Nombre_Archivo) {
	 
		 ArrayList<Hoja> Salida=new ArrayList<Hoja>();
		 
	  try {
	 
	   /**
	 
	    * Crea una nueva instancia de la clase FileInputStream
	 
	    */
	 
	   FileInputStream fileInputStream = new FileInputStream(
	 
	     Nombre_Archivo);
	 
	   /**
	 
	    * Crea una nueva instancia de la clase XSSFWorkBook
	 
	    */
	 
	   XSSFWorkbook Libro_trabajo = new XSSFWorkbook(fileInputStream);
	 
	   
	   
	   int NStilos=Libro_trabajo.getNumberOfSheets();
	   int columProcess=0;
		 
	   for (int i = 0; i < NStilos; i++) {
		   XSSFSheet Hoja_hssf = Libro_trabajo.getSheetAt(i);
		   HojaNueva Hojax=new HojaNueva(Hoja_hssf.getSheetName());
		   
		   Iterator<Row> Iterador_de_Fila = Hoja_hssf.rowIterator();
			 
		   List<List<XSSFCell>> Lista_Datos_Celda2 = new ArrayList<List<XSSFCell>>();
		   
		   boolean primera=true;
		   while (Iterador_de_Fila.hasNext()) {
		 
			   XSSFRow Fila_hssf = (XSSFRow) Iterador_de_Fila.next();
			   
			   
			   List<XSSFCell> Lista=new LinkedList<>();
			   
		 
			   if (primera)
			    {
				   Iterator<Cell> iterador = Fila_hssf.cellIterator();
			    	while (iterador.hasNext()) {
			    		XSSFCell Celda_hssf = (XSSFCell) iterador.next();
			    		Lista.add(Celda_hssf);
					 
					    }
			    	
			    	columProcess=Lista.size();
			    	
			    	primera=false;
			    }
			    else
			    {
			    	for (int j = 0; j < columProcess; j++) {
			    		XSSFCell Celda_hssf=Fila_hssf.getCell(j);
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
	 
	 private ArrayList<Hoja> GENERAR_XLS(String Nombre_Archivo) {
	 
		 ArrayList<Hoja> Salida=new ArrayList<Hoja>();
		 
	  try {
	 
	   /**
	 
	    * Crea una nueva instancia de la clase FileInputStream
	 
	    */
	 
	   FileInputStream fileInputStream = new FileInputStream(
	 
	     Nombre_Archivo);
	 
	   /**
	 
	    * Crea una nueva instancia de la clase POIFSFileSystem
	 
	    */
	 
	   POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
	 
	   /**
	 
	    * Crea una nueva instancia de la clase HSSFWorkBook
	 
	    */
	 
	   HSSFWorkbook Libro_trabajo = new HSSFWorkbook(fsFileSystem);
	   
	   int NStilos=Libro_trabajo.getNumberOfSheets();
	   int columProcess=0;
	   
	   for (int i = 0; i < NStilos; i++) {
		   HSSFSheet Hoja_hssf = Libro_trabajo.getSheetAt(i);
		   HojaAntigua Hojax=new HojaAntigua(Hoja_hssf.getSheetName());
		   
		   Iterator<Row> Iterador_de_Fila = Hoja_hssf.rowIterator();
			 
		   List<List<HSSFCell>> Lista_Datos_Celda2 = new ArrayList<List<HSSFCell>>();
		   
		   boolean primera=true;
		   
		   while (Iterador_de_Fila.hasNext()) {
		 
		    HSSFRow Fila_hssf = (HSSFRow) Iterador_de_Fila.next();
		 
		    List<HSSFCell> Lista=new LinkedList<>();
		    
		    if (primera)
		    {
		    	 Iterator<Cell> iterador = Fila_hssf.cellIterator();
		    	while (iterador.hasNext()) {
		    		 HSSFCell Celda_hssf = (HSSFCell) iterador.next();
		    		Lista.add(Celda_hssf);
				 
				    }
		    	
		    	columProcess=Lista.size();
		    	primera=false;
		    }
		    else
		    {
		    	for (int j = 0; j < columProcess; j++) {
		    		HSSFCell Celda_hssf=Fila_hssf.getCell(j);
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
	 
	  *
	 
	  * @param Datos_celdas
	 
	  *            - Listado de los datos que hay en la hoja de cálculo.
	 
	  */
	 
	 private void Imprimir_Consola(List<Hoja> HojasEntrada) {
	 
		 
	
	for (Hoja hoja : HojasEntrada) {
		
//		System.out.println("Nombre: " + hoja.getName());
		
		CompleteGrammar Grammar=new CompleteGrammar(hoja.getName(), hoja.getName(), coleccionstatica);
		coleccionstatica.getMetamodelGrammar().add(Grammar);
		HashMap<Integer, CompleteTextElementType> Hash=new HashMap<Integer, CompleteTextElementType>();
		HashMap<String, CompleteTextElementType> HashPath=new HashMap<String, CompleteTextElementType>();
		
		 CompleteTextElementType Descripccion=null;
		 CompleteTextElementType Icon=null;
		
		
		
		if (hoja instanceof HojaAntigua)
		{
			
			List<List<HSSFCell>> Datos_celdas = ((HojaAntigua) hoja).getListaHijos();
			
			 String Valor_de_celda;
			 
			
			 
			  for (int i = 0; i < Datos_celdas.size(); i++) {
			 
				CompleteDocuments Doc=new CompleteDocuments(coleccionstatica, Integer.toString(i), "");  
				if (i!=0)
					coleccionstatica.getEstructuras().add(Doc);
				  
			   List<HSSFCell> Lista_celda_temporal = Datos_celdas.get(i);
			 
			   for (int j = 0; j < Lista_celda_temporal.size(); j++) {
			 
			 
			     HSSFCell hssfCell = Lista_celda_temporal.get(j);
			 
			     
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
			    	
			    	CompleteTextElementType C=generaStructura(Valor_de_celda,Grammar,HashPath);
			    	Hash.put(new Integer(j), C);
//			    	System.out.print("Columna:" + Valor_de_celda + "\t\t");
			    	
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
//			    	System.out.print("Valor:" + Valor_de_celda + "\t\t");
			    	
			    	if (C==Descripccion)
			    		Doc.setDescriptionText(Valor_de_celda);
			    	
			    	if (C==Icon)
			    		Doc.setIcon(Valor_de_celda);
			    	
			    	}
				   }
			   }
			 
//			   System.out.println();
			 
			  }
		}
		else if (hoja instanceof HojaNueva)
		{
			
			
			
			List<List<XSSFCell>> Datos_celdas = ((HojaNueva) hoja).getListaHijos();
			
			 String Valor_de_celda;
			 
			  for (int i = 0; i < Datos_celdas.size(); i++) {
			 
				  List<XSSFCell> Lista_celda_temporal = Datos_celdas.get(i);
				  
				  CompleteDocuments Doc=new CompleteDocuments(coleccionstatica, Integer.toString(i), "");  
					if (i!=0)
						coleccionstatica.getEstructuras().add(Doc);
			 
			   for (int j = 0; j < Lista_celda_temporal.size(); j++) {
			 
				   Valor_de_celda="";
			 
				   if (Lista_celda_temporal.get(j)!=null)
				   {
			     XSSFCell hssfCell = (XSSFCell) Lista_celda_temporal.get(j);
			 
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
			    	CompleteTextElementType C=generaStructura(Valor_de_celda,Grammar,HashPath);
			    	Hash.put(new Integer(j), C);
//			    	System.out.print("Columna:" + Valor_de_celda + "\t\t");
			    	
			    	
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
//			    	System.out.print("Valor:" + Valor_de_celda + "\t\t");
			    	
			    	if (C==Descripccion)
			    		Doc.setDescriptionText(Valor_de_celda);
			    	
			    	if (C==Icon)
			    		Doc.setIcon(Valor_de_celda);
			    	
			    	}
			   }
			   }
			 
			   System.out.println();
			 
			  }
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
	 
	  String fileName = "ejemplo2.xls";
	 
	  System.out.println(fileName);
	 
	 CollectionXLS C = new CollectionXLS();
	 C.Leer_Archivo_Excel(fileName);
	 
	 System.out.println(C.toString());
	 }

	@Override
	public void ProcessInstances() {
		
		
	}


	public CompleteCollection getColeccion() {
		return coleccionstatica;
	}
	
	
}
