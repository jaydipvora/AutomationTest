package AutoLogin;

public class ReadExcel {

	public  void read()throws Exception {
		ExcelUtils ex=new ExcelUtils();
		ex.setExcelFile(System.getProperty("user.dir")+"\\src\\AutoLogin\\Dataprovider.xls","Credential");
		
		System.out.println(ex.getCellData(1,0));
		System.out.println(ex.getCellData(1,1));

		System.out.println(ex.getCellData(2,0));
		System.out.println(ex.getCellData(2,1));

		System.out.println(ex.getCellData(3,0));
		System.out.println(ex.getCellData(3,1));
		}
	public static void Main(String[] args) throws Exception{
        System.out.println("In main class");
        ReadExcel a1 = new ReadExcel();
        a1.read();
    }
	
}
