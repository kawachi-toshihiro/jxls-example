package examples.jxls;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;

public class PoiHelper {

	private static final Log log = LogFactory.getLog(PoiHelper.class);

	/**
	 *
	 * 指定シートの指定範囲を結合します。
	 *
	 * @param sheet 
	 * @param row1
	 * @param row2
	 * @param col1
	 * @param col2
	 */
	public void mergeRegion(HSSFSheet sheet,int row1,int row2,int col1,int col2){
		CellRangeAddress range = new CellRangeAddress(row1, row2, col1, col2);
		sheet.addMergedRegion(range);
	}

	public void logDebug(String message){
		log.debug(message);
	}

	/**
	 * 指定セルのすぐ左のセルの値をクリアします。
	 *
	 *
	 * @param cell
	 */
	public void clearLeft(Cell cell){
		Cell target = cell.getRow().getCell(cell.getColumnIndex() -1);
		cell.getRow().removeCell(target);
	}

}
