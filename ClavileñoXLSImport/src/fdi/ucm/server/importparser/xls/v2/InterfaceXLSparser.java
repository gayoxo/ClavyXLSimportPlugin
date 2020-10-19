package fdi.ucm.server.importparser.xls.v2;

/**
 * Interface parseModel, funciones necesarias para parseal un objeto, parsear su modelo y sus instancias
 * @author Joaquin Gayoso-Cabada
 *
 */
public interface InterfaceXLSparser {

	/**
	 * Funcion inicial del proceso del modelo.
	 */
	public void ProcessAttributes();
	
	/**
	 * Funcion inicial del proceso las instancias.
	 */
	public void ProcessInstances();
	
}
