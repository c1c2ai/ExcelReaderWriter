package com.mycompany.excelreaderwriter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import javafx.scene.control.Cell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws FileNotFoundException, IOException
    {
    
    try {
     
    FileInputStream file = new FileInputStream(new File("/Users/chetanSminq/Downloads/Workbook1.xlsx"));
     
    //Get the workbook instance for XLS file 
    XSSFWorkbook workbook = new XSSFWorkbook(file);
 
    //Get first sheet from the workbook
    XSSFSheet sheet = workbook.getSheetAt(0);
     
    //Iterate through each rows from first sheet
    Iterator<Row> rowIterator = sheet.iterator();
    while(rowIterator.hasNext()) {
        Row row = rowIterator.next();
         
        //For each row, iterate through each columns
        Iterator<org.apache.poi.ss.usermodel.Cell> cellIterator = row.cellIterator();
        while(cellIterator.hasNext()) {
             
            org.apache.poi.ss.usermodel.Cell cell = cellIterator.next();
             
            
                    System.out.print(cell.getStringCellValue() + "\t\t");
                 
        }
        System.out.println("");
    }
    file.close();
    
    
    
    
    
    //writer
    
    XSSFWorkbook workbook1 = new XSSFWorkbook();
XSSFSheet sheet1 = workbook1.createSheet("Sample sheet");


Map<String, Object[]> data = new HashMap<String, Object[]>();
data.put("1", new Object[] {"Emp No.", "Name", "Salary"});
data.put("2", new Object[] {1d, "John", 1500000d});
data.put("3", new Object[] {2d, "Sam", 800000d});
data.put("4", new Object[] {3d, "Dean", 700000d});
 
Set<String> keyset = data.keySet();
int rownum = 0;
for (String key : keyset) {
    Row row = sheet1.createRow(rownum++);
    Object [] objArr = data.get(key);
    int cellnum = 0;
    for (Object obj : objArr) {
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(cellnum++);
        if(obj instanceof Date) 
            cell.setCellValue((Date)obj);
        else if(obj instanceof Boolean)
            cell.setCellValue((Boolean)obj);
        else if(obj instanceof String)
            cell.setCellValue((String)obj);
        else if(obj instanceof Double)
            cell.setCellValue((Double)obj);
    }
}


FileOutputStream out = 
        new FileOutputStream(new File("/Users/chetanSminq/Downloads/chetanDemo.xlsx"));
workbook1.write(out);        
out.close();
System.out.println("Excel written successfully..");
    
} catch (FileNotFoundException e) {
    e.printStackTrace();
} catch (IOException e) {
    e.printStackTrace();
}

    }
}
