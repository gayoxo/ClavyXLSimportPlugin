/**
 * 
 */
package fdi.ucm.server.importparser.xls;

import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Date;
import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Clase que implementa la creacion de la base de datos per se
 * @author Joaquin Gayoso-Cabada
 *
 */
public class CollectionSQL implements InterfaceXLSparser {


	private static final String SQL_COLLECTION = null;
	private static final String COLECCION_A_APARTIR_DE_UN_SQL = null;
	private CompleteCollection coleccionstatica;
	
	public CollectionSQL() {
		coleccionstatica=new CompleteCollection(SQL_COLLECTION, COLECCION_A_APARTIR_DE_UN_SQL+ new Timestamp(new Date().getTime()));
	}
	
	/* (non-Javadoc)
	 * @see fdi.ucm.server.importparser.sql.SQLparserModel#ProcessAttributes()
	 */
	@Override
	public void ProcessAttributes() {
		
	}

	 boolean antiguo = false;
	 /**
	 
	  * Este metodo es usado para leer archivos Excel
	 
	  *
	 
	  * @param Nombre_Archivo
	 
	  *            - Nombre de archivo Excel.
	 
	  */
	 private void Leer_Archivo_Excel(String Nombre_Archivo) {
	 
	  /**
	 
	   * Crea una nueva instancia de Lista_Datos_Celda
	 
	   */
	 
	  List Lista_Datos_Celda = new ArrayList();
	 
	  if (Nombre_Archivo.contains(".xlsx")) {
	 
	   GENERAR_XLSX(Nombre_Archivo, Lista_Datos_Celda);
	 
	   antiguo = false;
	 
	  } else if (Nombre_Archivo.contains(".xls")) {
	 
	   GENERAR_XLS(Nombre_Archivo, Lista_Datos_Celda);
	 
	   antiguo = true;
	 
	  }
	 
	  /**
	 
	   * Llama el metodo Imprimir_Consola para imprimir los datos de la celda
	 
	   * en la consola.
	 
	   */
	 
	  Imprimir_Consola(Lista_Datos_Celda);
	 
	 }
	 
	 private void GENERAR_XLSX(String Nombre_Archivo, List Lista_Datos_Celda) {
	 
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
	 
	   XSSFSheet Hoja_hssf = Libro_trabajo.getSheetAt(0);
	 
	   /**
	 
	    * Iterar las filas y las celdas de la hoja de cálculo para obtener
	 
	    * toda la data.
	 
	    */
	 
	   Iterator Iterador_de_Fila = Hoja_hssf.rowIterator();
	 
	   while (Iterador_de_Fila.hasNext()) {
	 
	    XSSFRow Fila_hssf = (XSSFRow) Iterador_de_Fila.next();
	 
	    Iterator iterador = Fila_hssf.cellIterator();
	 
	    List Lista_celda_temporal = new ArrayList();
	 
	    while (iterador.hasNext()) {
	 
	     XSSFCell Celda_hssf = (XSSFCell) iterador.next();
	 
	     Lista_celda_temporal.add(Celda_hssf);
	 
	    }
	 
	    Lista_Datos_Celda.add(Lista_celda_temporal);
	 
	   }
	 
	  } catch (Exception e) {
	 
	   e.printStackTrace();
	 
	  }
	 
	 }
	 
	 private void GENERAR_XLS(String Nombre_Archivo, List Lista_Datos_Celda) {
	 
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
	 
	   HSSFSheet Hoja_hssf = Libro_trabajo.getSheetAt(0);
	 
	   /**
	 
	    * Iterar las filas y las celdas de la hoja de cálculo para obtener
	 
	    * toda la data.
	 
	    */
	 
	   Iterator Iterador_de_Fila = Hoja_hssf.rowIterator();
	 
	   while (Iterador_de_Fila.hasNext()) {
	 
	    HSSFRow Fila_hssf = (HSSFRow) Iterador_de_Fila.next();
	 
	    Iterator iterador = Fila_hssf.cellIterator();
	 
	    List Lista_celda_temporal = new ArrayList();
	 
	    while (iterador.hasNext()) {
	 
	     HSSFCell Celda_hssf = (HSSFCell) iterador.next();
	 
	     Lista_celda_temporal.add(Celda_hssf);
	 
	    }
	 
	    Lista_Datos_Celda.add(Lista_celda_temporal);
	 
	   }
	 
	  } catch (Exception e) {
	 
	   e.printStackTrace();
	 
	  }
	 
	 }
	 
	 /**
	 
	  * Este método se utiliza para imprimir los datos de la celda a la consola.
	 
	  *
	 
	  * @param Datos_celdas
	 
	  *            - Listado de los datos que hay en la hoja de cálculo.
	 
	  */
	 
	 private void Imprimir_Consola(List Datos_celdas) {
	 
	  String Valor_de_celda;
	 
	  for (int i = 0; i < Datos_celdas.size(); i++) {
	 
	   List Lista_celda_temporal = (List) Datos_celdas.get(i);
	 
	   for (int j = 0; j < Lista_celda_temporal.size(); j++) {
	 
	    if (antiguo) {
	 
	     HSSFCell hssfCell = (HSSFCell) Lista_celda_temporal.get(j);
	 
	     Valor_de_celda = hssfCell.toString();
	 
	    } else {
	 
	     XSSFCell hssfCell = (XSSFCell) Lista_celda_temporal.get(j);
	 
	     Valor_de_celda = hssfCell.toString();
	 
	    }
	 
	    System.out.print(Valor_de_celda + "\t");
	 
	   }
	 
	   System.out.println();
	 
	  }
	 
	 }
	 
	 public static void main(String[] args) {
	 
	  String fileName = "ejemplo2.xlsx";
	 
	  System.out.println(fileName);
	 
	  new CollectionSQL().Leer_Archivo_Excel(fileName);
	 
	 }

	@Override
	public void ProcessInstances() {
		// TODO Auto-generated method stub
		
	}


	public CompleteCollection getColeccion() {
		return coleccionstatica;
	}
	
	
}
