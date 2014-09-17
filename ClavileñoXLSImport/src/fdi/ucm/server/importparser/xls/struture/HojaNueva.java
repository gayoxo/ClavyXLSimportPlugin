/**
 * 
 */
package fdi.ucm.server.importparser.xls.struture;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;


/**
 * @author Joaquin Gayoso-Cabada
 *
 */
public class HojaNueva extends Hoja {

	private java.util.List<List<XSSFCell>> ListaHijos;

	public HojaNueva(String name) {
		super(name);
		ListaHijos=new ArrayList<List<XSSFCell>>();
	}

	/**
	 * @return the listaHijos
	 */
	public java.util.List<List<XSSFCell>> getListaHijos() {
		return ListaHijos;
	}

	/**
	 * @param listaHijos the listaHijos to set
	 */
	public void setListaHijos(java.util.List<List<XSSFCell>> listaHijos) {
		ListaHijos = listaHijos;
	}
}
