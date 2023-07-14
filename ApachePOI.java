package ApachePOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApachePOI {

	public static Map<String, List<List<String>>> readWorkbook() throws IOException{
		
		List<List<String>> alllist = new ArrayList<>();
		Map<String,List<List<String>>> data= new HashMap< String,List<List<String>>>();
		
		FileInputStream fis = new FileInputStream(new File(System.getProperty("user.dir")
				+ "\\src\\test\\java\\ApachePOI\\TestData.xlsx"));
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis); 		//workbook
		Iterator<Sheet> itrsheet = workbook.iterator();		//Sheet iterator
		while(itrsheet.hasNext()) {							//while loop for sheet iterator
			XSSFSheet sheet = (XSSFSheet) itrsheet.next();	//sheet
		Iterator<Row> itrrow =sheet.iterator();				//row iteraator
		  
		while(itrrow.hasNext()) {							//while loop for rows
			
			Row row = itrrow.next();
			Iterator<Cell> itrcell = row.iterator();
			List<String> list = new ArrayList<>();			//list to store all cell values
			
			while(itrcell.hasNext()) {						//while loop for single row
				Cell cell = itrcell.next();
			//CellType celltype = cell.getCellType();	
				
			list.add(cell.getStringCellValue());			//add cell value to list
				
			}
			 	
			
			alllist.add(list);								//add cell value lists to alllist list
			
		}
		 
		data.put(sheet.getSheetName(), alllist); 			//add to Map	
		}
		return data;
		
}

}
