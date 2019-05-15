package File.Excel;

import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import xsv.XSSFCell;

public class App {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		
		
		 ArrayList arList=null;
		    ArrayList al=null;
		    String fName = "C:\\Users\\sabrmc\\Downloads\\run.csv";
		    String thisLine; 
		    int count=0; 
		     FileInputStream fis = new FileInputStream(fName);
		     DataInputStream myInput = new DataInputStream(fis);
		    int z=0;
		    arList = new ArrayList();
		    while ((thisLine = myInput.readLine()) != null)
		    {
		     al = new ArrayList();
		     String strar[] = thisLine.split(",");
		     for(int j=0;j<strar.length;j++)
		     {
		     al.add(strar[j]);
		     }
		     arList.add(al);
		     System.out.println();
		     z++;
		    } 
		 
		    try
		    {
		     XSSFWorkbook hwb = new XSSFWorkbook();
		     XSSFSheet sheet = hwb.createSheet("new sheet");
		      for(int k=0;k<arList.size();k++)
		      {
		       ArrayList ardata = (ArrayList)arList.get(k);
		       XSSFRow row = sheet.createRow((short) 0+k);
		       for(int p=0;p<ardata.size();p++)
		       {
		        org.apache.poi.xssf.usermodel.XSSFCell cell = row.createCell((short) p);
		        String data = ardata.get(p).toString();
		        if(data.startsWith("=")){
		         cell.setCellType(Cell.CELL_TYPE_STRING);
		         data=data.replaceAll("\"", "");
		         data=data.replaceAll("=", "");
		         cell.setCellValue(data);
		        }else if(data.startsWith("\"")){
		            data=data.replaceAll("\"", "");
		            cell.setCellType(Cell.CELL_TYPE_STRING);
		            cell.setCellValue(data);
		        }else{
		            data=data.replaceAll("\"", "");
		            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		            cell.setCellValue(data);
		        }
		        //
		     //   cell.setCellValue(ardata.get(p).toString());
		       }
		       System.out.println();
		      } 
		     FileOutputStream fileOut = new FileOutputStream("run.xlsx");
		     hwb.write(fileOut);
		     fileOut.close();
		     System.out.println("Your excel file has been generated");
		    } catch ( Exception ex ) {
		         ex.printStackTrace();
		    }
		    
		    
		    ArrayList arList1=null;
		    ArrayList al1=null;
		    String fName1 = "C:\\Users\\sabrmc\\Downloads\\run(1).csv";
		    String thisLine1; 
		    int count1=0; 
		     FileInputStream fis1 = new FileInputStream(fName1);
		     DataInputStream myInput1 = new DataInputStream(fis1);
		    int z1=0;
		    arList1 = new ArrayList();
		    while ((thisLine1 = myInput1.readLine()) != null)
		    {
		     al1 = new ArrayList();
		     String strar[] = thisLine1.split(",");
		     for(int j=0;j<strar.length;j++)
		     {
		     al1.add(strar[j]);
		     }
		     arList1.add(al1);
		     System.out.println();
		     z1++;
		    } 
		 
		    try
		    {
		     XSSFWorkbook hwb = new XSSFWorkbook();
		     XSSFSheet sheet = hwb.createSheet("new sheet");
		      for(int k=0;k<arList1.size();k++)
		      {
		       ArrayList ardata = (ArrayList)arList1.get(k);
		       XSSFRow row = sheet.createRow((short) 0+k);
		       for(int p=0;p<ardata.size();p++)
		       {
		        org.apache.poi.xssf.usermodel.XSSFCell cell = row.createCell((short) p);
		        String data = ardata.get(p).toString();
		        if(data.startsWith("=")){
		         cell.setCellType(Cell.CELL_TYPE_STRING);
		         data=data.replaceAll("\"", "");
		         data=data.replaceAll("=", "");
		         cell.setCellValue(data);
		        }else if(data.startsWith("\"")){
		            data=data.replaceAll("\"", "");
		            cell.setCellType(Cell.CELL_TYPE_STRING);
		            cell.setCellValue(data);
		        }else{
		            data=data.replaceAll("\"", "");
		            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		            cell.setCellValue(data);
		        }
		        //*/
		     //   cell.setCellValue(ardata.get(p).toString());
		       }
		       System.out.println();
		      } 
		     FileOutputStream fileOut = new FileOutputStream("run(1).xlsx");
		     hwb.write(fileOut);
		     fileOut.close();
		     System.out.println("Your excel file has been generated");
		    } catch ( Exception ex ) {
		         ex.printStackTrace();
		    }
		    
		    
		    
		    ArrayList arList2=null;
		    ArrayList al2=null;
		    String fName2 = "C:\\Users\\sabrmc\\Downloads\\run(2).csv";
		    String thisLine2; 
		    int count2=0; 
		     FileInputStream fis2 = new FileInputStream(fName2);
		     DataInputStream myInput2 = new DataInputStream(fis2);
		    int z2=0;
		    arList2 = new ArrayList();
		    while ((thisLine2 = myInput2.readLine()) != null)
		    {
		     al2 = new ArrayList();
		     String strar[] = thisLine2.split(",");
		     for(int j=0;j<strar.length;j++)
		     {
		     al2.add(strar[j]);
		     }
		     arList2.add(al2);
		     System.out.println();
		     z2++;
		    } 
		 
		    try
		    {
		     XSSFWorkbook hwb = new XSSFWorkbook();
		     XSSFSheet sheet = hwb.createSheet("new sheet");
		      for(int k=0;k<arList2.size();k++)
		      {
		       ArrayList ardata = (ArrayList)arList2.get(k);
		       XSSFRow row = sheet.createRow((short) 0+k);
		       for(int p=0;p<ardata.size();p++)
		       {
		        org.apache.poi.xssf.usermodel.XSSFCell cell = row.createCell((short) p);
		        String data = ardata.get(p).toString();
		        if(data.startsWith("=")){
		         cell.setCellType(Cell.CELL_TYPE_STRING);
		         data=data.replaceAll("\"", "");
		         data=data.replaceAll("=", "");
		         cell.setCellValue(data);
		        }else if(data.startsWith("\"")){
		            data=data.replaceAll("\"", "");
		            cell.setCellType(Cell.CELL_TYPE_STRING);
		            cell.setCellValue(data);
		        }else{
		            data=data.replaceAll("\"", "");
		            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		            cell.setCellValue(data);
		        }
		        //*/
		     //   cell.setCellValue(ardata.get(p).toString());
		       }
		       System.out.println();
		      } 
		     FileOutputStream fileOut = new FileOutputStream("run(2).xlsx");
		     hwb.write(fileOut);
		     fileOut.close();
		     System.out.println("Your excel file has been generated");
		    } catch ( Exception ex ) {
		         ex.printStackTrace();
		    }
		    
		    
		    
		    ArrayList arList3=null;
		    ArrayList al3=null;
		    String fName3 = "C:\\Users\\sabrmc\\Downloads\\run(3).csv";
		    String thisLine3; 
		    int count3=0; 
		     FileInputStream fis3 = new FileInputStream(fName3);
		     DataInputStream myInput3 = new DataInputStream(fis3);
		    int z3=0;
		    arList3 = new ArrayList();
		    while ((thisLine3 = myInput3.readLine()) != null)
		    {
		     al3 = new ArrayList();
		     String strar[] = thisLine3.split(",");
		     for(int j=0;j<strar.length;j++)
		     {
		     al3.add(strar[j]);
		     }
		     arList3.add(al3);
		     System.out.println();
		     z3++;
		    } 
		 
		    try
		    {
		     XSSFWorkbook hwb = new XSSFWorkbook();
		     XSSFSheet sheet = hwb.createSheet("new sheet");
		      for(int k=0;k<arList3.size();k++)
		      {
		       ArrayList ardata = (ArrayList)arList3.get(k);
		       XSSFRow row = sheet.createRow((short) 0+k);
		       for(int p=0;p<ardata.size();p++)
		       {
		        org.apache.poi.xssf.usermodel.XSSFCell cell = row.createCell((short) p);
		        String data = ardata.get(p).toString();
		        if(data.startsWith("=")){
		         cell.setCellType(Cell.CELL_TYPE_STRING);
		         data=data.replaceAll("\"", "");
		         data=data.replaceAll("=", "");
		         cell.setCellValue(data);
		        }else if(data.startsWith("\"")){
		            data=data.replaceAll("\"", "");
		            cell.setCellType(Cell.CELL_TYPE_STRING);
		            cell.setCellValue(data);
		        }else{
		            data=data.replaceAll("\"", "");
		            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		            cell.setCellValue(data);
		        }
		        //*/
		     //   cell.setCellValue(ardata.get(p).toString());
		       }
		       System.out.println();
		      } 
		     FileOutputStream fileOut = new FileOutputStream("run(3).xlsx");
		     hwb.write(fileOut);
		     fileOut.close();
		     System.out.println("Your excel file has been generated");
		    } catch ( Exception ex ) {
		         ex.printStackTrace();
		    }
		    
		
		

    	XSSFWorkbook xSSFWorkbook=new XSSFWorkbook(new FileInputStream("run.xlsx"));
		XSSFSheet sheet=xSSFWorkbook.getSheetAt(0);
		System.out.println(sheet.getLastRowNum());
		XSSFRow[] row=new XSSFRow[3];
		
		ArrayList<RabbitModel> arrayList=new ArrayList<RabbitModel>();
		
		for(int i=0;i<=2;i++)
		{
			RabbitModel rm=new RabbitModel();
			row[i]=sheet.getRow(i+9);
			ArrayList<String> list=new ArrayList<String>();		
			if (i==0)
			{
				list.add(row[i].getCell(0).getStringCellValue());
			}
			else if(i==1) 
			{
				list.add(row[i].getCell(0).getStringCellValue());
			}
			
			else if(i==2)
			{
				list.add(row[i].getCell(0).getStringCellValue());
			}
			for (int j=2;j<=12;j++)
			{
				list.add(row[i].getCell(j).getStringCellValue());
			}
			rm.setList(list);
			arrayList.add(rm);	
		}
		
		for (RabbitModel rabbitModel2 : arrayList) {
			System.out.println(rabbitModel2.getList());
		}
		//return arrayList;	
		
		XSSFWorkbook xSSFWorkbook2=new XSSFWorkbook(new FileInputStream("run(1).xlsx"));
		XSSFSheet sheet2=xSSFWorkbook2.getSheetAt(0);
		System.out.println(sheet2.getLastRowNum());
		XSSFRow[] row2=new XSSFRow[3];
		
		ArrayList<RabbitModel> arrayList2=new ArrayList<RabbitModel>();
		
		for(int i=0;i<=2;i++)
		{
			RabbitModel rm=new RabbitModel();
			row2[i]=sheet2.getRow(i+9);
			ArrayList<String> list=new ArrayList<String>();		
			if (i==0)
			{
				list.add(row2[i].getCell(0).getStringCellValue());
			}
			else if(i==1) 
			{
				list.add(row2[i].getCell(0).getStringCellValue());
			}
			
			else if(i==2)
			{
				list.add(row2[i].getCell(0).getStringCellValue());
			}
			for (int j=2;j<=12;j++)
			{
				list.add(row2[i].getCell(j).getStringCellValue());
				System.out.println(row2[i].getCell(j).getStringCellValue());
			}
			rm.setList(list);
			arrayList2.add(rm);	
		}
		
		
		//System.out.println(arrayList2);
		
		for (RabbitModel rabbitModel2 : arrayList) {
			System.out.println(rabbitModel2.getList());
		}
		
		
		XSSFWorkbook xSSFWorkbook3=new XSSFWorkbook(new FileInputStream("run(2).xlsx"));
		XSSFSheet sheet3=xSSFWorkbook3.getSheetAt(0);
		System.out.println(sheet3.getLastRowNum());
		XSSFRow[] row3=new XSSFRow[3];
		
		ArrayList<RabbitModel> arrayList3=new ArrayList<RabbitModel>();
		
		for(int i=0;i<=2;i++)
		{
			RabbitModel rm=new RabbitModel();
			row3[i]=sheet3.getRow(i+9);
			ArrayList<String> list=new ArrayList<String>();
			
			if (i==0)
			{
				list.add(row3[i].getCell(0).getStringCellValue());
			}
			else if(i==1) 
			{
				list.add(row3[i].getCell(0).getStringCellValue());
			}
			
			else if(i==2)
			{
				list.add(row3[i].getCell(0).getStringCellValue());
			}
			
			for (int j=2;j<=12;j++)
			{
				list.add(row3[i].getCell(j).getStringCellValue());
			}
			rm.setList(list);
			arrayList3.add(rm);	
		}
		
		
		XSSFWorkbook xSSFWorkbook4=new XSSFWorkbook(new FileInputStream("run(3).xlsx"));
		XSSFSheet sheet4=xSSFWorkbook4.getSheetAt(0);
		System.out.println(sheet4.getLastRowNum());
		XSSFRow[] row4=new XSSFRow[3];
		
		ArrayList<RabbitModel> arrayList4=new ArrayList<RabbitModel>();
		
		for(int i=0;i<=2;i++)
		{
			RabbitModel rm=new RabbitModel();
			row4[i]=sheet4.getRow(i+9);
			ArrayList<String> list=new ArrayList<String>();
			
			if (i==0)
			{
				list.add(row4[i].getCell(0).getStringCellValue());
			}
			else if(i==1) 
			{
				list.add(row4[i].getCell(0).getStringCellValue());
			}
			
			else if(i==2)
			{
				list.add(row4[i].getCell(0).getStringCellValue());
			}
			
			for (int j=2;j<=12;j++)
			{
				list.add(row4[i].getCell(j).getStringCellValue());
			}
			rm.setList(list);
			arrayList4.add(rm);	
		}
		
			
		SimpleDateFormat formatter = new SimpleDateFormat("MM/dd/yyyy");
		SimpleDateFormat formatter2 = new SimpleDateFormat("EEEE");
		Date date = new Date();
	    System.out.println(formatter.format(date));
	    System.out.println(formatter2.format(date));
			
		Workbook workbook0 = new XSSFWorkbook();
		Sheet sheet0=workbook0.createSheet("op");
		
		
		int rowNum = 0;
				
		Row rowdate = sheet0.createRow(rowNum++);
		Row rowday = sheet0.createRow(rowNum++);
		Cell cell10=rowdate.createCell(0);
	    cell10.setCellValue("BIS_JBOSS_WS_PROD");
	    Cell cell20=rowday.createCell(0);
	    cell20.setCellValue("");
	    
	    
	    CellStyle style10 = workbook0.createCellStyle();
    	style10.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
        style10.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
    	cell10.setCellStyle(style10);
    	
    	CellStyle style20 = workbook0.createCellStyle();
    	style20.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
        style20.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
    	cell20.setCellStyle(style20);
		
		for (int i=1,j=10;i<=11;j--)
		{
			Calendar calendar = Calendar.getInstance();
	        calendar.setTime(date);
	        calendar.add(Calendar.DAY_OF_YEAR, (-j));
	        Date previousDate = calendar.getTime();
	        //System.out.println(formatter.format(previousDate));		
			Cell cell=rowdate.createCell(i);
			cell.setCellValue(formatter.format(previousDate));
			
			CellStyle style = workbook0.createCellStyle();
        	style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
            style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        	cell.setCellStyle(style);
			
			Cell cell1=rowday.createCell(i++);
			cell1.setCellValue(formatter2.format(previousDate));
			
			CellStyle style1 = workbook0.createCellStyle();
        	style1.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
            style1.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        	cell1.setCellStyle(style1);
			
			System.out.print(rowdate.getCell(i-1).getStringCellValue());
		}
		System.out.println();
		
        for(RabbitModel rabbitModel : arrayList) {
            Row row0 = sheet0.createRow(rowNum++);
            int colNum=0;
            System.out.println();
            
            for (String list: rabbitModel.list) {
            	//System.out.println(list);
            	Cell cell=row0.createCell(colNum++);
            	cell.setCellValue(new XSSFRichTextString(list));
            	
            	if(rowNum!=4)
            	{
            	CellStyle style = workbook0.createCellStyle();
            	style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
                style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
            	cell.setCellStyle(style);
            	System.out.print(row0.getCell(colNum-1).getStringCellValue());
            	}
			}
           }
        System.out.println();
        
        
        Row rowdate2 = sheet0.createRow(rowNum++);
		Row rowday2 = sheet0.createRow(rowNum++);	
		Cell cell11=rowdate2.createCell(0);
		cell11.setCellValue("BIS_DOTNET_WS_PROD");
		Cell cell21=rowday2.createCell(0);
	    cell21.setCellValue("");
		
    	cell11.setCellStyle(style10);   	
    	cell21.setCellStyle(style20);
		
		for (int i=1,j=10;i<=11;j--)
		{
			Calendar calendar = Calendar.getInstance();
	        calendar.setTime(date);
	        calendar.add(Calendar.DAY_OF_YEAR, (-j));
	        Date previousDate = calendar.getTime();
	        //System.out.println(formatter.format(previousDate));
	        Cell cell=rowdate2.createCell(i);
			cell.setCellValue(formatter.format(previousDate));
			
			CellStyle style = workbook0.createCellStyle();
        	style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
            style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        	cell.setCellStyle(style);
			
			Cell cell1=rowday2.createCell(i++);
			cell1.setCellValue(formatter2.format(previousDate));
			
			CellStyle style1 = workbook0.createCellStyle();
        	style1.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
            style1.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        	cell1.setCellStyle(style1);
			System.out.print(rowdate.getCell(i-1).getStringCellValue());
		}
		System.out.println();
               
        for(RabbitModel rabbitModel : arrayList2) {
            Row row0 = sheet0.createRow(rowNum++);
            int colNum=0;
            System.out.println();

            for (String list: rabbitModel.list) {
            	//System.out.println(list);
            	Cell cell=row0.createCell(colNum++);
            	cell.setCellValue(new XSSFRichTextString(list));
            	
            	if(rowNum!=9)
            	{
            	CellStyle style = workbook0.createCellStyle();
            	style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
                style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
            	cell.setCellStyle(style);
            	System.out.print(row0.getCell(colNum-1).getStringCellValue());
            	}
			}
           }
        
               
        Row rowdate3 = sheet0.createRow(rowNum++);
		Row rowday3 = sheet0.createRow(rowNum++);	
		Cell cell12=rowdate3.createCell(0);
		cell12.setCellValue("CAAPI_INTERNAL");
		Cell cell22=rowday3.createCell(0);
	    cell22.setCellValue("");
		
    	cell12.setCellStyle(style10);    	
    	cell22.setCellStyle(style20);
		
		for (int i=1,j=10;i<=11;j--)
		{
			Calendar calendar = Calendar.getInstance();
	        calendar.setTime(date);
	        calendar.add(Calendar.DAY_OF_YEAR, (-j));
	        Date previousDate = calendar.getTime();
	        //System.out.println(formatter.format(previousDate));
	        Cell cell=rowdate3.createCell(i);
			cell.setCellValue(formatter.format(previousDate));
			
			CellStyle style = workbook0.createCellStyle();
        	style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
            style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        	cell.setCellStyle(style);
			
			Cell cell1=rowday3.createCell(i++);
			cell1.setCellValue(formatter2.format(previousDate));
			
			CellStyle style1 = workbook0.createCellStyle();
        	style1.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
            style1.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        	cell1.setCellStyle(style1);
			System.out.print(rowdate3.getCell(i-1).getStringCellValue());
		}
		System.out.println();
               
        for(RabbitModel rabbitModel : arrayList3) {
            Row row0 = sheet0.createRow(rowNum++);
            int colNum=0;
            System.out.println();

            for (String list: rabbitModel.list) {
            	//System.out.println(list);
            	Cell cell=row0.createCell(colNum++);
            	cell.setCellValue(new XSSFRichTextString(list));
            	
            	if(rowNum!=14)
            	{
            	CellStyle style = workbook0.createCellStyle();
            	style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
                style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
            	cell.setCellStyle(style);
            	System.out.print(row0.getCell(colNum-1).getStringCellValue());
            	}
			}
           }
        
        
        Row rowdate4 = sheet0.createRow(rowNum++);
		Row rowday4 = sheet0.createRow(rowNum++);	
		Cell cell13=rowdate4.createCell(0);
		cell13.setCellValue("CAAPI_EXTERNAL");
		
		Cell cell23=rowday4.createCell(0);
	    cell23.setCellValue("");

    	cell13.setCellStyle(style10);
    	cell23.setCellStyle(style20);
		
		for (int i=1,j=10;i<=11;j--)
		{
			Calendar calendar = Calendar.getInstance();
	        calendar.setTime(date);
	        calendar.add(Calendar.DAY_OF_YEAR, (-j));
	        Date previousDate = calendar.getTime();
	        //System.out.println(formatter.format(previousDate));
	        Cell cell=rowdate4.createCell(i);
			cell.setCellValue(formatter.format(previousDate));
			
			CellStyle style = workbook0.createCellStyle();
        	style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
            style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        	cell.setCellStyle(style);
			
			Cell cell1=rowday4.createCell(i++);
			cell1.setCellValue(formatter2.format(previousDate));
			
			CellStyle style1 = workbook0.createCellStyle();
        	style1.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
            style1.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        	cell1.setCellStyle(style1);
			System.out.print(rowdate4.getCell(i-1).getStringCellValue());
		}
		System.out.println();
               
        for(RabbitModel rabbitModel : arrayList) {
            Row row0 = sheet0.createRow(rowNum++);
            int colNum=0;
            System.out.println();

            for (String list: rabbitModel.list) {
            	//System.out.println(list);
            	Cell cell=row0.createCell(colNum++);
            	cell.setCellValue(new XSSFRichTextString(list));
            	
            	if(rowNum!=19)
            	{
            	CellStyle style = workbook0.createCellStyle();
            	style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
                style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
            	cell.setCellStyle(style);
            	System.out.print(row0.getCell(colNum-1).getStringCellValue());
            	}
			}
             }
                 
        for(int i = 0; i<=11; i++) {
            sheet0.autoSizeColumn(i);
        }
                
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
        workbook0.write(fileOut);
        fileOut.close();
        //Closing the workbook
        workbook0.close();
	}
}
