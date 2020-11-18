/**
 * 
 */
package fdi.ucm.server.importparser.xls.v3.struture;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;


/**
 * @author Joaquin Gayoso-Cabada
 *
 */
public class HojaV2 extends Hoja {

	private java.util.List<List<Cell>> ListaHijos;

	public HojaV2(String name) {
		super(name);
		ListaHijos=new ArrayList<List<Cell>>();
	}

	/**
	 * @return the listaHijos
	 */
	public java.util.List<List<Cell>> getListaHijos() {
		return ListaHijos;
	}

	/**
	 * @param listaHijos the listaHijos to set
	 */
	public void setListaHijos(java.util.List<List<Cell>> listaHijos) {
		ListaHijos = listaHijos;
	}
}
