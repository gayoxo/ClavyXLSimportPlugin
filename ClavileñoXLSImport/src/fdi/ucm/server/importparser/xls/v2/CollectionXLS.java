/**
 * 
 */
package fdi.ucm.server.importparser.xls.v2;

import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;

import fdi.ucm.server.importparser.xls.v2.struture.Hoja;
import fdi.ucm.server.importparser.xls.v2.struture.HojaV2;
import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteLinkElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteTextElement;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteLinkElementType;
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

import org.apache.lucene.analysis.standard.StandardAnalyzer;
import org.apache.lucene.document.Document;
import org.apache.lucene.document.LongField;
import org.apache.lucene.document.TextField;
import org.apache.lucene.document.Field;
import org.apache.lucene.index.DirectoryReader;
import org.apache.lucene.index.IndexReader;
import org.apache.lucene.index.IndexWriter;
import org.apache.lucene.index.IndexWriterConfig;
import org.apache.lucene.index.Term;
import org.apache.lucene.search.BooleanQuery;
import org.apache.lucene.search.FuzzyQuery;
import org.apache.lucene.search.IndexSearcher;
import org.apache.lucene.search.Query;
import org.apache.lucene.search.ScoreDoc;
import org.apache.lucene.search.TopDocs;
import org.apache.lucene.store.Directory;
import org.apache.lucene.store.RAMDirectory;
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
	
	 class Dest_Por
	 {
		 Long Destino=0l;
		 Long Porcentaje=100l;

		 public Dest_Por(Long destino,Long porcentaje) {
			 Destino=destino;
			 Porcentaje=porcentaje;
		 }

		public Long getDestino() {
			return Destino;
		}

		public void setDestino(Long destino) {
			Destino = destino;
		}

		public Long getPorcentaje() {
			return Porcentaje;
		}

		public void setPorcentaje(Long porcentaje) {
			Porcentaje = porcentaje;
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
		 

		 
		 
		 
		 HashSet<Long> Destinos=new HashSet<Long>();
		
		 
		 HashMap<String, HashMap<Integer,Dest_Por>> LinkElements=new HashMap<String, HashMap<Integer,Dest_Por>>();
		 HashMap<Long, Long> ListSyme=new HashMap<Long, Long>();
		 
		 try {

			 JSONParser parser = new JSONParser();

			 Object obj = parser.parse(new FileReader(fileNameP));
					 
			 JSONArray jsonObject = (JSONArray) obj;
			 
			 for (int i = 0; i < jsonObject.size(); i++) {
				JSONObject array_element = (JSONObject) jsonObject.get(i);
				
				try {
					String Gram=array_element.get("g").toString();
					String COL_S=array_element.get("c").toString();
					Long DES_ID=(Long) array_element.get("d");
					Long FuzzyO=(Long) array_element.get("s");
					
					Long FuzzyI=100l;
					if (FuzzyO!=null)
						FuzzyI=Long.parseLong(FuzzyO.toString());
					
					if (Gram!=null&&COL_S!=null&&Gram.trim().length()>0&&COL_S.trim().length()>0&&DES_ID>0&&FuzzyI>0)
						{

						
						Integer I = calculaColumnaIntArray(COL_S.toUpperCase());
						
						HashMap<Integer, Dest_Por> listaPre= LinkElements.get(Gram);
						if (listaPre==null)
							listaPre=new HashMap<Integer, Dest_Por>();
						
						
							Dest_Por valorD=new Dest_Por(DES_ID, FuzzyI);	
							
						listaPre.put(I,valorD);
						LinkElements.put(Gram, listaPre);
						
						Destinos.add(DES_ID);
						
						}
				} catch (Exception e) {
					e.printStackTrace();
				}
				
				
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		 
		 procesaGramarHash(preCol.getMetamodelGrammar(),ListSyme,Destinos);
		 
	 
		 try {
			 StandardAnalyzer analyzer = new StandardAnalyzer();
			 Directory index = new RAMDirectory();

			 IndexWriterConfig config = new IndexWriterConfig(analyzer);

			 IndexWriter w = new IndexWriter(index, config);
			
			 HashMap<Long, CompleteDocuments> ListaHash=new HashMap<Long, CompleteDocuments>();
			 
			 
			 for (CompleteDocuments docu : preCol.getEstructuras()) {
				Document doc = new Document();
				doc.add(new LongField("CID", docu.getClavilenoid(), Field.Store.YES));
				for (CompleteElement elem : docu.getDescription()) {
					if (ListSyme.containsKey(elem.getHastype().getClavilenoid())&&elem instanceof CompleteTextElement)
						doc.add(new TextField("C"+ListSyme.get(elem.getHastype().getClavilenoid()),
								((CompleteTextElement)elem).getValue(), Field.Store.YES));
				}
				w.addDocument(doc);
				ListaHash.put(docu.getClavilenoid(), docu);
			}
			 
			 ArrayList<Hoja> Hojas=new ArrayList<Hoja>();

			 w.close();
			 
			  if (Nombre_Archivo.contains(".xlsx")) 
				  Hojas=GENERAR(Nombre_Archivo,FileFormat.NEW); 
			 else if (Nombre_Archivo.contains(".xls")) 
				  Hojas=GENERAR(Nombre_Archivo,FileFormat.OLD);

			  Imprimir_Consola(Hojas,LinkElements,index,analyzer,Destinos,ListaHash,log,Nombre_Archivo,ColDesc);
			 
			
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

	private void procesaGramarHash(List<CompleteGrammar> metamodelGrammar,
			 HashMap<Long, Long> listSyme, HashSet<Long> destinos) {
		for (CompleteGrammar completeGrammar : metamodelGrammar) 
			procesaStructureHash(completeGrammar.getSons(),listSyme,destinos);

		
	}

	private void procesaStructureHash(List<CompleteElementType> sons, HashMap<Long, Long> listSyme,
			HashSet<Long> destinos) {
		for (CompleteElementType revisa : sons) {
			if (destinos.contains(revisa.getClassOfIterator().getClavilenoid()))
			{
			//valido
				
				listSyme.put(revisa.getClavilenoid(),revisa.getClassOfIterator().getClavilenoid());
				
				CompleteElementType revisaB = revisa.getBSon();
				while (revisaB!=null)
				{
					listSyme.put(revisa.getClavilenoid(),revisaB.getClassOfIterator().getClavilenoid());
					revisaB = revisaB.getBSon();
				}
				
				
				
			}
			
			procesaStructureHash(revisa.getSons(), listSyme, destinos);
		}
		
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
			 HashMap<String, HashMap<Integer,Dest_Por>> linkElements,
			 Directory index, StandardAnalyzer analyzer, HashSet<Long> destinos,
			 HashMap<Long, CompleteDocuments> listaDocumento, List<String> LogsOut,
			 String Filename, Integer colDesc) throws IOException {
	 
		 
	
	for (Hoja hoja : HojasEntrada) {
		
//		System.out.println("Nombre: " + hoja.getName());
		
		
		CompleteGrammar Grammar=new CompleteGrammar(hoja.getName(), Filename, coleccionstatica);
		coleccionstatica.getMetamodelGrammar().add(Grammar);
		HashMap<Integer, CompleteTextElementType> Hash=new HashMap<Integer, CompleteTextElementType>();
		HashMap<Integer, CompleteLinkElementType> HashL=new HashMap<Integer, CompleteLinkElementType>();
		HashMap<String, CompleteTextElementType> HashPath=new HashMap<String, CompleteTextElementType>();
		
		 CompleteTextElementType Descripccion=null;
		 CompleteTextElementType Icon=null;
		
		 List<List<Cell>> Datos_celdas = ((HojaV2) hoja).getListaHijos();
		 
	        // 3. search
	        int hitsPerPage = 10;
	        IndexReader reader = DirectoryReader.open(index);
	        IndexSearcher searcher = new IndexSearcher(reader);

			
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
			    	
			    	
			    	
			    	CompleteTextElementType C=generaStructura(Valor_de_celda,Grammar,HashPath);
			    	Hash.put(new Integer(j), C);
//			    	System.out.print("Columna:" + Valor_de_celda + "\t\t");
			    	
			    	if (colDesc==j||isDesscription(Valor_de_celda))
			    		Descripccion=C;
			    	
			    	if (isIcon(Valor_de_celda))
			    		Icon=C;



				   }
			   
		   }	
		  
		  
		  HashMap<Long, CompleteLinkElementType> listaNewLinks=new HashMap<Long, CompleteLinkElementType>();
		  
		  for (int j = 0; j < Lista_celda_temporal0.size(); j++) 
		  if (linkElements.get(hoja.getName()) != null&&linkElements.get(hoja.getName()).containsKey(new Integer(j)))
	    		{
			  CompleteTextElementType C=Hash.get(new Integer(j));
			  
			  Dest_Por destinosLoca = linkElements.get(hoja.getName()).get(new Integer(j));
			  
			  if (listaNewLinks.get(destinosLoca.getDestino())!=null)
			  {
				  HashL.put(new Integer(j), listaNewLinks.get(destinosLoca.getDestino()));
			  }
			  else
			  {
			  
			  CompleteLinkElementType CL=new CompleteLinkElementType("LINKED",C.getFather(), C.getCollectionFather());
			  CL.setSons(C.getSons());
			  
			  if (C.getFather()==null)
				  C.getCollectionFather().getSons().add(CL);
			  else
				  C.getFather().getSons().add(CL);

			  listaNewLinks.put(destinosLoca.getDestino(), CL);
			  
//			  Hash.remove(new Integer(j));
			  HashL.put(new Integer(j), CL);
			  System.out.println(j+":"+CL.getName()+"->> Deberia ser un Link");
			  }
			  
			 
	    	
	    		}
			 
			 
		  HashMap<String,List<CompleteLinkElement>> NuevasPersonas=new  HashMap<String,List<CompleteLinkElement>>();
		  
		  CompleteDocuments DocAUX=new CompleteDocuments(null, "aux", "aux");  
		  
		  
		  for (int i = 1; i < Datos_celdas.size(); i++) {
		 
			CompleteDocuments Doc=new CompleteDocuments(coleccionstatica, Integer.toString(i), "");  
			if (i!=0)
				coleccionstatica.getEstructuras().add(Doc);
			  
		   List<Cell> Lista_celda_temporal = Datos_celdas.get(i);
		 
			  HashMap<CompleteLinkElementType, List<String>> searchPatern = new HashMap<CompleteLinkElementType, List<String>>();

		   
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
		 
		 

		    	CompleteTextElementType C=Hash.get(new Integer(j));
		    	

		    	
		    	if (C!=null)
		    	{
		    	CompleteTextElement CT=new CompleteTextElement(C, Valor_de_celda);
		    	Doc.getDescription().add(CT);

		    	
		    	if (C==Descripccion)
		    		Doc.setDescriptionText(Valor_de_celda);
		    	
		    	if (C==Icon)
		    		Doc.setIcon(Valor_de_celda);
		    	}

		    		CompleteLinkElementType C2=HashL.get(new Integer(j));
		    		if (C2!=null)
			    	{
		    			
		    			
		    			List<String> listaA=searchPatern.get(C2);
		    			
		    			if (listaA==null)
		    				listaA=new ArrayList<String>();
		    			
		    			listaA.add(Valor_de_celda);
		    			
		    			searchPatern.put(C2, listaA);
		    			
		    			
		    			
		    			
		    	        
		    			
		    	
			    	
		    	}
		    	
			   }
			   
			   
		   }
		   
		 //AQUIN
		   
		   for (Entry<CompleteLinkElementType, List<String>> C2_val : searchPatern.entrySet()) {
			   BooleanQuery.Builder builder = new BooleanQuery.Builder();
				
				for (Long long1 : destinos) {
					
					for (String valor : C2_val.getValue()) {
						Query q0 = new FuzzyQuery(new Term("C"+long1, valor));
						builder.add(q0, 
								org.apache.lucene.search.BooleanClause.Occur.SHOULD);
					}
					
					
				}

				
				
				StringBuffer nameCOmplete=new StringBuffer();
				boolean af=true;
				for (String long1 : C2_val.getValue()) 
					{
					
					if (af)
						af=false;
					else
						nameCOmplete.append(", ");
						
					nameCOmplete.append(long1);
					
					}
				

				
				Query q = builder.build();
				
				


		        TopDocs docs = searcher.search(q, hitsPerPage);
		        
		        
		        
		        ScoreDoc[] hits = docs.scoreDocs;
		        
		       // System.out.println("Found " + hits.length + " hits. for ->"+ Arrays.toString(C2_val.getValue().toArray()));
		        
		        if (hits.length>0)
		        {
		        	 int docId = hits[0].doc;
		        	 Document d = searcher.doc(docId);
		        	 Long LL=Long.parseLong(d.get("CID"));
		        	 
		        	 if (listaDocumento.get(LL)!=null)
		        	 {
		        		CompleteDocuments DocuLin = listaDocumento.get(LL); 
		        	 CompleteLinkElement CT=new CompleteLinkElement(C2_val.getKey(), listaDocumento.get(LL));
			    	Doc.getDescription().add(CT);

			    	LogsOut.add("LINKED:"+nameCOmplete.toString()+" LINKED to ->> D:"+DocuLin.getClavilenoid()+" "+DocuLin.getDescriptionText());
			    	
		        	 }
		        }else
		        {
		        	if (C2_val.getValue().size()==1)
			        	try {
			        		
			        		Long LL = Long.parseLong(C2_val.getValue().get(0));
			        		 if (listaDocumento.get(LL)!=null)
		   	        	 {
		   	        	 CompleteLinkElement CT=new CompleteLinkElement(C2_val.getKey(), listaDocumento.get(LL));
					    	Doc.getDescription().add(CT);
	
		   	        	 }
						} catch (Exception e) {
							
							 CompleteLinkElement CT=new CompleteLinkElement(C2_val.getKey(), DocAUX);
						    	Doc.getDescription().add(CT);
						    	
							List<CompleteLinkElement> ListaElemdoc = NuevasPersonas.get(nameCOmplete.toString());
							if (ListaElemdoc==null)
								ListaElemdoc=new ArrayList<CompleteLinkElement>();
							
							ListaElemdoc.add(CT);
							
							NuevasPersonas.put(nameCOmplete.toString(), ListaElemdoc);
							
						}
		        	else
		        		{
		        		 CompleteLinkElement CT=new CompleteLinkElement(C2_val.getKey(), DocAUX);
					    	Doc.getDescription().add(CT);
					    	
						List<CompleteLinkElement> ListaElemdoc = NuevasPersonas.get(nameCOmplete.toString());
						if (ListaElemdoc==null)
							ListaElemdoc=new ArrayList<CompleteLinkElement>();
						
						ListaElemdoc.add(CT);
						
						NuevasPersonas.put(nameCOmplete.toString(), ListaElemdoc);
		        		}
		        }
		}
		   

		 
		  }
		 
		  if (NuevasPersonas.size()>0)
		  {
			  CompleteGrammar GrammarPP=new CompleteGrammar("CREATED", Filename, coleccionstatica);
				coleccionstatica.getMetamodelGrammar().add(GrammarPP);
				
				CompleteTextElementType Nombre =new CompleteTextElementType("VALUE", GrammarPP);
				GrammarPP.getSons().add(Nombre);
				
		  for (Entry<String, List<CompleteLinkElement>> doccrea : NuevasPersonas.entrySet()) {
			  LogsOut.add("CREATED:"+doccrea.getKey()+" CREATED");
			  
			  CompleteDocuments Doc=new CompleteDocuments(coleccionstatica, doccrea.getKey(), "");  
			  coleccionstatica.getEstructuras().add(Doc);
			  CompleteTextElement tete=new CompleteTextElement(Nombre, doccrea.getKey());
			  Doc.getDescription().add(tete);
			  
			  for (CompleteLinkElement linkvaria : doccrea.getValue()) {
				  linkvaria.setValue(Doc);
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
