package excelReadExample;

import java.io.IOException;

public class ExcelRead {

	public static void main(String[] args) throws IOException {
		ExcelReadExample obj=new ExcelReadExample();
		String s=obj.readData(1, 0);
		System.out.println(s);
		obj.readFunction();
		

	}

}
