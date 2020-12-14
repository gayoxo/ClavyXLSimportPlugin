/**
 * 
 */
package fdi.ucm.server.importparser.xls.v3;

import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;

import fdi.ucm.server.importparser.xls.v3.struture.Hoja;
import fdi.ucm.server.importparser.xls.v3.struture.HojaV2;
import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteLinkElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteTextElement;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteLinkElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteOperationalValueType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteTextElementType;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map.Entry;
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
	
	 
	 
	 class ColV_ColId
	 {
		 Integer ColumnaValor=0;
		 Integer ColumnaId=100;

		 public ColV_ColId(Integer ColumnaId,Integer ColumnaValor) {
			 this.ColumnaValor=ColumnaValor;
			 this.ColumnaId=ColumnaId;
		 }

		public Integer getColumnaValor() {
			return ColumnaValor;
		}

		public void setColumnaValor(Integer columnaValor) {
			ColumnaValor = columnaValor;
		}

		public Integer getColumnaId() {
			return ColumnaId;
		}

		public void setColumnaId(Integer columnaId) {
			ColumnaId = columnaId;
		}

		
		 
		 
		 
	 }
	
	
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
	 * @param preCol 
	 * @param log 
	 * @param columns 
	 
	  */
	 public void Leer_Archivo_Excel(String Nombre_Archivo, String fileNameP,
			 CompleteCollection preCol, List<String> log, String columns) {
	 
	  /**
	 
	   * Crea una nueva instancia de Lista_Datos_Celda
	 
	   */
		 Integer ColDesc=-1;
		 
		 if (columns!=null)
			 ColDesc=calculaColumnaIntArray(columns);
		 

		 
		 
		
		 HashMap<String, HashSet<ColV_ColId>> LinkElements=new HashMap<String, HashSet<ColV_ColId>>();
		 
		 try {

			 JSONParser parser = new JSONParser();

			 Object obj = parser.parse(new FileReader(fileNameP));
					 
			 JSONArray jsonObject = (JSONArray) obj;
			 
			 for (int i = 0; i < jsonObject.size(); i++) {
				JSONObject array_element = (JSONObject) jsonObject.get(i);
				
				try {
					String sheet=array_element.get("Sheet").toString();
					String ColumnaValor=array_element.get("ColVal").toString();
					String ColId=array_element.get("ColId").toString();
					

					
					if (sheet!=null&&ColumnaValor!=null&&ColId!=null&&!sheet.isEmpty()&&!ColumnaValor.isEmpty()&&!ColId.isEmpty())
						{

						
						Integer IT = calculaColumnaIntArray(ColumnaValor.toUpperCase());
						Integer ID = calculaColumnaIntArray(ColId.toUpperCase());
						
						HashSet<ColV_ColId> listaPre= LinkElements.get(sheet);
						if (listaPre==null)
							listaPre=new HashSet<ColV_ColId>();
						
						
							ColV_ColId valorD=new ColV_ColId(ID, IT);	
							
						listaPre.add(valorD);
						LinkElements.put(sheet, listaPre);
						
						
						}
				} catch (Exception e) {
					e.printStackTrace();
				}
				
				
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		 

	 
		 try {

			 HashMap<Long, CompleteDocuments> ListaHash=new HashMap<Long, CompleteDocuments>();
			 
			 
			 for (CompleteDocuments docu : preCol.getEstructuras())				
				{
				 for (CompleteElement elementosDoc : docu.getDescription()) {
					CompleteElementType CET=elementosDoc.getHastype();
					for (CompleteOperationalValueType covt : CET.getShows()) {
						if (covt.getName().equals("oldid")&&elementosDoc instanceof CompleteTextElement)
						{
							try {
								Long idold=Long.parseLong(((CompleteTextElement)elementosDoc).getValue());
								ListaHash.put(idold, docu);
							} catch (Exception e) {
								// no hay problema de que falle
							}
						}
					}
				}
				 
				 ListaHash.put(docu.getClavilenoid(), docu);
				}
			
			 
			 ArrayList<Hoja> Hojas=new ArrayList<Hoja>();

			
			 
			  if (Nombre_Archivo.contains(".xlsx")) 
				  Hojas=GENERAR(Nombre_Archivo,FileFormat.NEW); 
			 else if (Nombre_Archivo.contains(".xls")) 
				  Hojas=GENERAR(Nombre_Archivo,FileFormat.OLD);

			  Imprimir_Consola(Hojas,LinkElements,ListaHash,log,Nombre_Archivo,ColDesc);
			 
			
		} catch (Exception e) {
			e.printStackTrace();
			log.add("Error Index Cross");
			return;
		}
	
	
		 
		 
		 
	 
	 
	 }
	 
	 private Integer calculaColumnaIntArray(String columns) {
		return calculaColumnaIntAbs(columns)-1;
	}

	private int calculaColumnaIntAbs(String columnsIn) {
		 Integer Final=0;
		    
		 StringBuilder input1 = new StringBuilder();
		 
		 input1.append(columnsIn);
		 
		 input1=input1.reverse();
		 
		 String columns = input1.toString();
		 
		    int BaseZ = 26;
		    
		    for (int j=0; j<columns.toUpperCase().length();j++)
		    {
		        int pos = columns.toUpperCase().length()-1-j;
		        char charar = columns.toUpperCase().charAt(pos);
		        
		        Integer I = ((int)(new Character(charar))) - ((int)(new Character('A')))+1;
		        
		        
		       double Var = Math.pow(BaseZ, pos)*I;

		        Final=Final+(int)Var;
		    }
		    
		    return Final;
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
	 * @param index 
	 * @param analyzer 
	 * @param destinos 
	 * @param colDesc 
	 * @param preCol,List<StringS> LogsOut 
	 * @param listaHash 
	 
	  *
	 
	  * @param Datos_celdas
	 
	  *            - Listado de los datos que hay en la hoja de cálculo.
	 * @throws IOException 
	 
	  */
	 
	 private void Imprimir_Consola(List<Hoja> HojasEntrada,
			 HashMap<String, HashSet<ColV_ColId>> linkElements,
			 HashMap<Long, CompleteDocuments> listaDocumento, List<String> LogsOut,
			 String Filename, Integer colDesc) throws IOException {
	 
		 
	
	for (Hoja hoja : HojasEntrada) {
		
//		System.out.println("Nombre: " + hoja.getName());
		
		
		CompleteGrammar Grammar=new CompleteGrammar(hoja.getName(), Filename, coleccionstatica);
		coleccionstatica.getMetamodelGrammar().add(Grammar);
		HashMap<Integer, CompleteElementType> Hash=new HashMap<Integer, CompleteElementType>();
		HashMap<String, CompleteElementType> HashPath=new HashMap<String, CompleteElementType>();
		
		 CompleteTextElementType Descripccion=null;
		 CompleteTextElementType Icon=null;
		
		 List<List<Cell>> Datos_celdas = ((HojaV2) hoja).getListaHijos();
		 

		 HashMap<Integer, List<Integer>> equivalenciaLinks_col_valoires=new HashMap<Integer, List<Integer>>();
		  
		  if (linkElements.get(hoja.getName()) != null) {
			  HashSet<ColV_ColId> tablavalores = linkElements.get(hoja.getName());
			  for (ColV_ColId val_id : tablavalores) {
				  
				  if (equivalenciaLinks_col_valoires.get(val_id.getColumnaId())==null)
					  equivalenciaLinks_col_valoires.put(val_id.getColumnaId(),new LinkedList<Integer>());
				  
				  equivalenciaLinks_col_valoires.get(val_id.getColumnaId()).add(val_id.getColumnaValor());

				  
			}
		  }
		 
			
		 String Valor_de_celda;

			 //AQUI LA PRIMERA
		   List<Cell> Lista_celda_temporal0 = Datos_celdas.get(0);
		 
		  for (int j = 0; j < Lista_celda_temporal0.size(); j++) {
			   
			   Cell hssfCell = Lista_celda_temporal0.get(j);
				 
			     
			     Valor_de_celda="";
				 
				   if (Lista_celda_temporal0.get(j)!=null)
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
			 
			 
			    	if (Valor_de_celda==null||Valor_de_celda.isEmpty())
			    		Valor_de_celda=hoja.getName()+" Columna:"+j;
			    	
			    	
			    	
			    	CompleteElementType C=generaStructura(Valor_de_celda,Grammar,HashPath);
			    	
			    	
			    	Hash.put(new Integer(j), C);
//			    	System.out.print("Columna:" + Valor_de_celda + "\t\t");
			    	
			    	if (colDesc==j||isDesscription(Valor_de_celda))
			    		Descripccion=(CompleteTextElementType)C;
			    	
			    	if (isIcon(Valor_de_celda))
			    		Icon=(CompleteTextElementType)C;



				   }
			   
		   }
		  
		  for (int j = 0; j < Lista_celda_temporal0.size(); j++) {
			  
			  if (equivalenciaLinks_col_valoires.get(new Integer(j))!=null)
			  {
				  CompleteElementType valor= Hash.get(new Integer(j));
				  
				  CompleteLinkElementType CLE=new CompleteLinkElementType(valor.getName(), valor.getFather(), valor.getCollectionFather());
				  CLE.setSons(valor.getSons());
				  
				  for (CompleteElementType HijoE : valor.getSons())
					HijoE.setFather(CLE);
				  
				  if (valor.getFather()!=null)
					  {
					  int pos=-1;
					  for (int i = 0; i < valor.getFather().getSons().size(); i++) 
						  if (valor.getFather().getSons().get(i)==valor)
							  pos=i;
						
					
					  if (pos==-1)
						  valor.getFather().getSons().add(CLE);
					  else
						  {
						  valor.getFather().getSons().remove(pos);
						  valor.getFather().getSons().add(pos,CLE);						  
						  }
					  }
				  else
					  {
					  
					  int pos=-1;
					  for (int i = 0; i < valor.getCollectionFather().getSons().size(); i++) 
						  if (valor.getCollectionFather().getSons().get(i)==valor)
							  pos=i;
						
					
					  if (pos==-1)
						  valor.getCollectionFather().getSons().add(CLE);
					  else
						  {
						  valor.getCollectionFather().getSons().remove(pos);
						  valor.getCollectionFather().getSons().add(pos,CLE);						  
						  }
					  
					  
					  
					  
					  }
				  
				  Hash.put(new Integer(j), CLE);
			  }
		  }
		  
		  
		  HashMap<Integer,HashMap<String,CompleteDocuments>> Nuevos=new  HashMap<Integer,HashMap<String, CompleteDocuments>>();
		  
		  for (int i = 1; i < Datos_celdas.size(); i++) {
		 
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
		 
			   }
			   

		    	CompleteElementType C=Hash.get(new Integer(j));
		    	

		    	
		    	if (C!=null)
		    	{
		    	
		    		CompleteElement CT=null;
		    	if (C instanceof CompleteTextElementType && !Valor_de_celda.isEmpty())	
		    		CT=new CompleteTextElement((CompleteTextElementType)C, Valor_de_celda);

		    	if (C instanceof CompleteLinkElementType)	
		    		{
		    		
		    		CompleteDocuments DocuLin=null;
						if (!Valor_de_celda.isEmpty())
						{
							
							try {
								Float LPV=Float.parseFloat(Valor_de_celda);	
								 DocuLin = listaDocumento.get(LPV.longValue()); 
								} catch (Exception e) {
									//NO PASA NADA
								}
							
							try {
							Long LPV=Long.parseLong(Valor_de_celda);	
							 DocuLin = listaDocumento.get(LPV); 
							} catch (Exception e) {
								//NO PASA NADA
							}
						}	
					
					if (DocuLin!=null)	
						{
						CT=new CompleteLinkElement((CompleteLinkElementType)C, DocuLin);
						 LogsOut.add("LINKED:"+DocuLin.getDescriptionText()+" LINKED OK");
						
						}
					else
						{
						List<Integer> Lj2 = equivalenciaLinks_col_valoires.get(new Integer(j));
						
						//Collections.sort(Lj2);
						
						StringBuffer SBValor=new StringBuffer();
						 
						
						for (Integer j2 : Lj2) {
							
						
					     
					     String Valor_de_celda2 = "";
						 
					     Cell hssfCell2 = Lista_celda_temporal.get(j2);
						 
						   if (Lista_celda_temporal.get(j2)!=null)
						   {
					     
					     
					     if(hssfCell2.getCellType() == Cell.CELL_TYPE_FORMULA){
					    	 switch(hssfCell2.getCachedFormulaResultType()) {
					            case Cell.CELL_TYPE_NUMERIC:
					                System.out.println("Last evaluated as: " + hssfCell2.getNumericCellValue());
					                Valor_de_celda2=Double.toString(hssfCell.getNumericCellValue());
					                break;
					            case Cell.CELL_TYPE_STRING:
					                System.out.println("Last evaluated as \"" + hssfCell2.getRichStringCellValue() + "\"");
					                Valor_de_celda2=hssfCell2.getRichStringCellValue().toString();
					                break;
					             default:
					            	 Valor_de_celda2 = hssfCell2.toString();
							        break;
					                	
					        }
					     }else
					    	 Valor_de_celda2 = hssfCell2.toString();
					     
					     SBValor.append(Valor_de_celda2+" ");
						   }
					     
						}
						
						HashMap<String, CompleteDocuments> NuevosTabla = Nuevos.get(new Integer(j));
						if (NuevosTabla==null)
							NuevosTabla=new HashMap<String, CompleteDocuments>();
						
						
					     CompleteDocuments nuevoYaCreado=NuevosTabla.get(SBValor.toString().trim());
						if (nuevoYaCreado==null)
							{
							nuevoYaCreado=new CompleteDocuments(coleccionstatica, SBValor.toString().trim(), "");
							NuevosTabla.put(SBValor.toString().trim(), nuevoYaCreado);
							}
						
						CT=new CompleteLinkElement((CompleteLinkElementType)C, nuevoYaCreado);	
						
						
						Nuevos.put(new Integer(j), NuevosTabla);
						
		    		}
		    		}
		    	
		    	if (CT!=null)
		    		Doc.getDescription().add(CT);

		    	
		    	if (C==Descripccion)
		    		Doc.setDescriptionText(Valor_de_celda);
		    	
		    	if (C==Icon)
		    		Doc.setIcon(Valor_de_celda);
		    	


    	
		    	}
		    	
			   }
			   
			   
		   
		   
		
		   

		 
		  }
		 
		  if (Nuevos.size()>0)
		  {
			  
			  CompleteGrammar GrammarPP=new CompleteGrammar("CREATED", Filename, coleccionstatica);
				coleccionstatica.getMetamodelGrammar().add(GrammarPP);
			  
		  for (Entry<Integer, HashMap<String, CompleteDocuments>> doccrea_int : Nuevos.entrySet()) {
			  
			  CompleteElementType C=Hash.get(new Integer(doccrea_int.getKey()));
			  
			  String CName="~unamed";
			  if (C!=null)
				  CName=C.getName();
			  
			 
			
			CompleteTextElementType Nombre =new CompleteTextElementType("VALUE BY \""+ CName + "\"", GrammarPP);
			GrammarPP.getSons().add(Nombre);
			  
			  for (Entry<String, CompleteDocuments> doccrea : doccrea_int.getValue().entrySet()) {
				
			
			  LogsOut.add("CREATED BY \""+ CName + "\":"+doccrea.getKey()+" CREATED");
			  
			  CompleteDocuments Doc=doccrea.getValue();  
			  coleccionstatica.getEstructuras().add(Doc);
			  CompleteTextElement tete=new CompleteTextElement(Nombre, doccrea.getKey());
			  Doc.getDescription().add(tete);
			  
	        	
	        	
	        	
	        	
			}
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

	private CompleteElementType generaStructura(String valor_de_celda, CompleteGrammar grammar, HashMap<String, CompleteElementType> hashPath) {
		 
		
		 CompleteElementType preproducido = hashPath.get(valor_de_celda);
			if (preproducido!=null)
				return preproducido;
		 
		 
		String[] pathL=valor_de_celda.split("/");
		
		CompleteElementType Padre=null;
		
		 if (pathL.length>1)
			 Padre=producePadre(pathL,hashPath,grammar);
		 
		 CompleteElementType Salida=null;
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
			HashMap<String, CompleteElementType> hashPath,CompleteGrammar CG) {
		
		String Acumulado = "";
		CompleteElementType Padre = null;
		for (int i = 0; i < pathL.length-1; i++) {
			if (i!=0)
				Acumulado=Acumulado+"/"+pathL[i];
			else
				Acumulado=Acumulado+pathL[i];
			CompleteElementType yo = hashPath.get(Acumulado);
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
	 
		 String fileName = "testV3.xls";
		if (args.length>0)
			fileName=args[0];
		
		String fileNameP = "testV3.xls.json";
		if (args.length>1)
			fileNameP=args[1];
	 
		CompleteCollection fileNameCP = new CompleteCollection();
		if (args.length>2)
			{
			String fileNameCPS=args[2];
			try {
				 File file = new File(fileNameCPS);
				 FileInputStream fis = new FileInputStream(file);
				 ObjectInputStream ois = new ObjectInputStream(fis);
				 fileNameCP = (CompleteCollection) ois.readObject();
				 ois.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
			}
		
		
	  System.out.println(fileName);
	  
	  System.out.println(fileNameP);
	 
	  ArrayList<String> Logs = new ArrayList<String>();
	  
	 CollectionXLS C = new CollectionXLS();
	 C.Leer_Archivo_Excel(fileName,fileNameP,
			 fileNameCP,Logs,"C");
	 
	 System.out.println(C.toString());
	 
	 for (String string : Logs) {
		 System.out.println(string);
	}
	 
	 try {
			String FileIO = System.getProperty("user.home")+File.separator+System.currentTimeMillis()+".clavy";
			
			System.out.println(FileIO);
			
			ObjectOutputStream oos = new ObjectOutputStream(new FileOutputStream(FileIO));

			oos.writeObject(C.getColeccion());

			oos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	 
	 }

	@Override
	public void ProcessInstances() {
		
		
	}


	public CompleteCollection getColeccion() {
		return coleccionstatica;
	}
	
	
}
