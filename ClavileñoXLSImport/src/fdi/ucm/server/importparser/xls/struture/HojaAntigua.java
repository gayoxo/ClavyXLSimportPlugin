/**
 * 
 */
package fdi.ucm.server.importparser.xls.struture;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;


/**
 * @author Joaquin Gayoso-Cabada
 *Cobertura para XLS Antiguas
 *
 */
public class HojaAntigua extends Hoja{

	private java.util.List<List<HSSFCell>> ListaHijos;

	public HojaAntigua(String name) {
		super(name);
		ListaHijos=new ArrayList<List<HSSFCell>>();
	}

	/**
	 * @return the listaHijos
	 */
	public java.util.List<List<HSSFCell>> getListaHijos() {
		return ListaHijos;
	}

	/**
	 * @param listaHijos the listaHijos to set
	 */
	public void setListaHijos(java.util.List<List<HSSFCell>> listaHijos) {
		ListaHijos = listaHijos;
	}
	
	
}
