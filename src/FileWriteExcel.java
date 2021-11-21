import java.io.File;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;


public class FileWriteExcel {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		
//		Workbook object
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		
//		spreadsheet object
		
		HSSFSheet spreadsheet = workbook.createSheet("Person Data");
		
		HSSFRow row;
		
		Map<String, Object[]> personData = new TreeMap<String, Object[]>();
		
		personData.put("1",new Object[] {
				"ID", "Name","Email-ID","Address"
		});
		
		personData.put("2",new Object[] {
				"128", "Peter","peter@gmail.com","Quahog Rhode Island"
		});
		
		personData.put("3",new Object[] {
				"129", "Glen","glen@gmail.com","Quahog Rhode Island"
		});
		
		personData.put("4",new Object[] {
				"130", "Trevor","trevor130@gmail.com","LA"
		});
		
		personData.put("5",new Object[] {
				"131", "Joe Swanson","jeoswan@gmail.com","New York"
		});
		
		personData.put("6",new Object[] {
				"132", "Shubham","shubh132@gmail.com","Kalyan"
		});
		
		Set<String> keyid = personData.keySet();
		
		int rowid = 0;
		
		// writing the data into sheets...
		
		for(String key : keyid) {
			
			row = spreadsheet.createRow(rowid++);
			Object[] objectArr = personData.get(key);
			int cellid = 0;
			  
            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
            }
        }
		
		
	// .xlsx is the format for Excel Sheets...
    // writing the workbook into the file...
		try {
		 FileOutputStream out;
	
		 out = new FileOutputStream(new File("C:\\Users\\Akshay\\Desktop\\Springboot\\JavaExcel\\GFGSheets.xlsx"));
		 workbook.write(out);
		 out.close();
		 workbook.close();
		}
		catch(Exception ex) {
			ex.printStackTrace();
		}
	}

}

