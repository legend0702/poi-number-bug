package zhq;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 在Excel2007中 用700*0.35后产生的神奇现象
 */
public class POIReader {

	/**
	 * 获取第一个工作表的第二行第一个Cell(sheet1!A2) 并打印数值
	 * 
	 * @param args
	 * @throws FileNotFoundException
	 * @throws IOException
	 */

	public static void main(String[] args) throws FileNotFoundException,
			IOException {
		try (XSSFWorkbook wb = new XSSFWorkbook(
				ClassLoader.getSystemResourceAsStream("异常数字2016-03-30.xlsx"))) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(1);
			Cell cell = row.getCell(0);
			System.out.println(cell.getNumericCellValue());
		}
	}

}
