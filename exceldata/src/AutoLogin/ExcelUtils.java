package AutoLogin;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

	private  XSSFSheet ExcelWSheet;

	private  XSSFWorkbook ExcelWBook;

	public  final String Outpath = System.getProperty("user.dir")+"\\src\\AutoLogin\\Dataprovider.xls";

	public  void setExcelFile(String Path, String SheetName)
			throws Exception {

		try {
			File src = new File(Path);
			FileInputStream fis = new FileInputStream(src);
			ExcelWBook = new XSSFWorkbook(fis);
			ExcelWSheet = ExcelWBook.getSheet(SheetName);
		} catch (Exception e) {

			throw (e);

		}
	}

	public  String getCellData(int RowNum, int ColNum) throws Exception {

		try {
			String CellData = (ExcelWSheet.getRow(RowNum).getCell(ColNum)
					.getStringCellValue());

			return CellData;

		} catch (Exception e) {

			return "";

		}

	}

	public  void setCellData(String Result, int RowNum, int ColNum)
			throws Exception {

		try {

			ExcelWSheet.getRow(0).createCell(2).setCellValue(Result);
			FileOutputStream fileOut = new FileOutputStream(Outpath);

			ExcelWBook.write(fileOut);

			fileOut.flush();

			fileOut.close();

		} catch (Exception e) {

			throw (e);

		}

	}

}
