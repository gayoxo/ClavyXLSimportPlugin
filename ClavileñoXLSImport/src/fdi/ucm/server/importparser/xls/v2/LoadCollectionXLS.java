/**
 * 
 */
package fdi.ucm.server.importparser.xls.v2;

import java.util.ArrayList;
import java.util.Arrays;
import javax.management.RuntimeErrorException;

import fdi.ucm.server.modelComplete.ImportExportDataEnum;
import fdi.ucm.server.modelComplete.ImportExportPair;
import fdi.ucm.server.modelComplete.LoadCollection;
import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.CompleteCollectionAndLog;

/**
 * @author Joaquin Gayoso-Cabada
 *
 */
public class LoadCollectionXLS extends LoadCollection{

	private static ArrayList<ImportExportPair> Parametros;
	private CompleteCollection PreCol;
	
	
	public LoadCollectionXLS() {
		super();
	}
	
	@Override
	public CompleteCollectionAndLog processCollecccion(
			ArrayList<String> dateEntrada) {
		
		CollectionXLS C=null;
		 ArrayList<String> Log = new ArrayList<String>();
		if (dateEntrada.size()>0 && !dateEntrada.get(0).isEmpty())
		{ 
		String fileName = dateEntrada.get(0);
		
		String fileNameP="";
		if (dateEntrada.size()>1)
			fileNameP= dateEntrada.get(1);
		
		String Columns=null;
		if (dateEntrada.size()>2)
			{
			Columns= dateEntrada.get(2);
			}
		
		 System.out.println(fileName);
		 C = new CollectionXLS();


		 
		 
		 C.Leer_Archivo_Excel(fileName,fileNameP,PreCol,Log,Columns);
		}
		else
		{
			if (dateEntrada.size()<=0)
					Log.add("Error: Numero de Elementos de entrada invalidos");
			if (dateEntrada.get(0).isEmpty()) 
				Log.add("Error: Path del file vacio");
		}
		if (C.getColeccion()!=null)
		 return new CompleteCollectionAndLog(C.getColeccion(),Log);
		else
			throw new RuntimeErrorException(new Error(Arrays.toString(Log.toArray())));
	}

	@Override
	public ArrayList<ImportExportPair> getConfiguracion() {
		if (Parametros==null)
		{
			ArrayList<ImportExportPair> ListaCampos=new ArrayList<ImportExportPair>();
			ListaCampos.add(new ImportExportPair(ImportExportDataEnum.File, "Upload XLS File :"));
			ListaCampos.add(new ImportExportPair(ImportExportDataEnum.File, "Upload .json (Cross-Relation File) :",true));
			ListaCampos.add(new ImportExportPair(ImportExportDataEnum.Text, "Column letter in excel for description :",true));
			Parametros=ListaCampos;
			return ListaCampos;
		}
		else return Parametros;
	}

	@Override
	public String getName() {
		return "XLS V2";
	}

	@Override
	public boolean getCloneLocalFiles() {
		return false;
	}

	@Override
	public boolean needComplete() {
		return true;
	}
	
	@Override
	public void setcompleteCollectionPre(CompleteCollection pre) {
		PreCol=pre;
	}
}
