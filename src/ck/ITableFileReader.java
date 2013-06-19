package ck;

import java.util.Vector;

/**
 * Read table format file
 * @author shizexing
 *
 */
public interface ITableFileReader {
	/**
	 * Read data from file
	 */
	public void loadData();
	/**
	 * get data in memory
	 * @return
	 */
	public Vector<Vector<String>> getData();
}
