package sadagi.ericsson.softhuman.core.model;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xerces.impl.xpath.regex.ParseException;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import com.aspose.cells.ConsolidationFunction;
import com.aspose.cells.PivotFieldType;
import com.aspose.cells.PivotTable;
import com.aspose.cells.PivotTableAutoFormatType;
import com.aspose.cells.PivotTableCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
public class Testing {
	
	
	private static HashMap<String, String[]> HashMap_Fact ;
	private String[] ArrData_Fact;
	private LinkedList<String> length= new LinkedList<String>();
	private static ArrayList<Integer> Ar = new ArrayList<Integer>();

	//XSSFWorkbook wb = new XSSFWorkbook("C:/WORK/Final Input/Final Input/UPW 2G Worst Cell Tracker- Jul'16.xlsx");
	static String PATH ="C:/WORK/Final Input/Final Input/";
	static XSSFWorkbook wb = null;
	static XSSFSheet sheet = null;
	static String cell = null;
	public static void main(String[] args) throws IOException
	{	
		
		wb = new XSSFWorkbook(new FileInputStream(new File(PATH +"UPW 2G Worst Cell Tracker- Jul'16.xlsx")));
		 sheet = wb.getSheet("Sheet1");
		new Testing().output();
		try {
			new Testing().processOneSheet(PATH +"UPW 2G Worst Cell Tracker- Jul'16.xlsx", "Sheet1");
			
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 FileOutputStream outExcel=new FileOutputStream(new File(PATH +"UPW 2G Worst Cell Tracker- Jul'16.xlsx"));           
	        wb.write(outExcel); //write the output data
	        outExcel.close();   
		    wb.close();
	
	}
	public void output() throws IOException{
		
		try
		{
        BufferedReader br = new BufferedReader(new FileReader(PATH +"ERCS_UPW_BSS_BBH_REPORT_12072016.csv"));
        String currentLine=null;
        LinkedHashMap<String, Integer> HashMap_Column_Heading=new LinkedHashMap<String , Integer>();
        int count=0;
        String [] Array;
        String Site_ID="";
        Double TCH_Drop=0.0;
        Double TCH_BL=0.0;
        Double SD_Bl=0.0;
        Double SD_drop=0.0;
        Double SD_Assig=0.0;
        Double VFE_TASR=0.0;
        Double VFE_HSR=0.0;
        
        DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
        Date date2 = new Date();
        String date3=dateFormat.format(date2);

        XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		XSSFRow orow=null;
		XSSFCell ocell=null;
		orow=sheet.createRow(0);
		ocell=orow.createCell(0);
		ocell.setCellValue("Site_ID");
		ocell=orow.createCell(1);
		ocell.setCellValue("ERCS_BSS_TCH_DROP BH Value");
		ocell=orow.createCell(2);
		ocell.setCellValue("ERCS_BSS_Eric_TCH_BLOCKING BH Value");
		ocell=orow.createCell(3);
		ocell.setCellValue("ERCS_BSS_Eric_SDCCH_BLOCKING BH Value");
		ocell=orow.createCell(4);
		ocell.setCellValue("ERCS_BSS_SDCCH_Drop_Rate_EyeSpot BH Value");
		ocell=orow.createCell(5);
		ocell.setCellValue("ERCS_BSS_NQI_SD_Assign_Suc_Rate% BH Value");
		ocell=orow.createCell(6);
		ocell.setCellValue("ERCS_BSS_NQI_TCH_Assign_Suc_Rate% BH Value");
		ocell=orow.createCell(7);
		ocell.setCellValue("ERCS_BSS_NQI_HANDOVER_SUCCESS_RATE_% BH Value");
		ocell=orow.createCell(8);
		ocell.setCellValue("Date");
        while ((currentLine = br.readLine()) != null) 
        {
        	count++;
            if(count==1) 
            {
                 Array=currentLine.split(",");
                 for(int j=0;j<Array.length;j++)
                 {
                       HashMap_Column_Heading.put(Array[j], j);
                 }
                 
            }
            if(count>=2)
            {
            	Array=currentLine.split(",");
            	Site_ID="";
            	TCH_Drop=0.0;
            	TCH_BL=0.0;
            	SD_Bl=0.0;
            	SD_drop=0.0;
            	SD_Assig=0.0;
            	VFE_TASR=0.0;
            	VFE_HSR=0.0;
            	 orow=sheet.createRow(sheet.getPhysicalNumberOfRows());
            	
                if(Array[HashMap_Column_Heading.get("Short name")]!=null)//FETCHING THE COLUMN SHORT NAME FROM CVS FILE
                {
                	Site_ID=Array[HashMap_Column_Heading.get("Short name")];
                }
                length.add(Site_ID);
            	ocell=orow.createCell(0);
    			ocell.setCellValue(Site_ID);
               //System.out.println(Site_ID); 

                if(Array[HashMap_Column_Heading.get("ERCS_BSS_TCH_DROP BH Value")]!=null)
                {
                	if(!(Array[HashMap_Column_Heading.get("ERCS_BSS_TCH_DROP BH Value")].equals("")))
                	TCH_Drop=Double.parseDouble(Array[HashMap_Column_Heading.get("ERCS_BSS_TCH_DROP BH Value")]);
                	else
                		TCH_Drop=0.0;
                }
                ocell=orow.createCell(1);
    			ocell.setCellValue(TCH_Drop);
                
               // System.out.print(", " + TCH_Drop); 
                
                if(Array[HashMap_Column_Heading.get("ERCS_BSS_Eric_TCH_BLOCKING BH Value")]!=null)
                {
                	if(!(Array[HashMap_Column_Heading.get("ERCS_BSS_Eric_TCH_BLOCKING BH Value")].equals("")))
                	TCH_BL=Double.parseDouble(Array[HashMap_Column_Heading.get("ERCS_BSS_Eric_TCH_BLOCKING BH Value")]);
                	else
                		TCH_BL=0.0;
                }
                ocell=orow.createCell(2);
    			ocell.setCellValue(TCH_BL);
               // System.out.println(", " + TCH_BL);
                
                if(Array[HashMap_Column_Heading.get("ERCS_BSS_Eric_SDCCH_BLOCKING BH Value")]!=null)
                {
                	if(!(Array[HashMap_Column_Heading.get("ERCS_BSS_Eric_SDCCH_BLOCKING BH Value")].equals("")))
                	SD_Bl=Double.parseDouble(Array[HashMap_Column_Heading.get("ERCS_BSS_Eric_SDCCH_BLOCKING BH Value")]);
                	else
                		SD_Bl=0.0;
                }
                ocell=orow.createCell(3);
    			ocell.setCellValue(SD_Bl);
                
                //System.out.println(", " + SD_Bl);
                
                if(Array[HashMap_Column_Heading.get("ERCS_BSS_SDCCH_Drop_Rate_EyeSpot BH Value")]!=null)
                {
                	if(!(Array[HashMap_Column_Heading.get("ERCS_BSS_SDCCH_Drop_Rate_EyeSpot BH Value")].equals("")))
                	SD_drop=Double.parseDouble(Array[HashMap_Column_Heading.get("ERCS_BSS_SDCCH_Drop_Rate_EyeSpot BH Value")]);
                	else
                		SD_drop=0.0;
                }
                ocell=orow.createCell(4);
    			ocell.setCellValue(SD_drop);
                
    			
    			ocell=orow.createCell(8);
    			ocell.setCellValue(date3);
               // System.out.println(", " + SD_drop);
                
                if(Array[HashMap_Column_Heading.get("ERCS_BSS_NQI_SD_Assign_Suc_Rate% BH Value")]!=null)
                {
                	if(!(Array[HashMap_Column_Heading.get("ERCS_BSS_NQI_SD_Assign_Suc_Rate% BH Value")].equals("")))
                	SD_Assig=Double.parseDouble(Array[HashMap_Column_Heading.get("ERCS_BSS_NQI_SD_Assign_Suc_Rate% BH Value")]);
                	else
                		SD_Assig=0.0;
                }
                ocell=orow.createCell(5);
    			ocell.setCellValue(SD_Assig);
                //System.out.println(", " + SD_Assig);
                
                
                if(Array[HashMap_Column_Heading.get("ERCS_BSS_NQI_TCH_Assign_Suc_Rate% BH Value")]!=null)
                {
                	if(!(Array[HashMap_Column_Heading.get("ERCS_BSS_NQI_TCH_Assign_Suc_Rate% BH Value")].equals("")))
                	VFE_TASR=Double.parseDouble(Array[HashMap_Column_Heading.get("ERCS_BSS_NQI_TCH_Assign_Suc_Rate% BH Value")]);
                	else
                		VFE_TASR=0.0;
                }
                ocell=orow.createCell(6);
    			ocell.setCellValue(VFE_TASR);
               // System.out.println(", " + VFE_TASR);
                
                if(Array[HashMap_Column_Heading.get("ERCS_BSS_NQI_HANDOVER_SUCCESS_RATE_% BH Value")]!=null)
                {
                    if(!(Array[HashMap_Column_Heading.get("ERCS_BSS_NQI_HANDOVER_SUCCESS_RATE_% BH Value")].equals("")))
                	VFE_HSR=Double.parseDouble(Array[HashMap_Column_Heading.get("ERCS_BSS_NQI_HANDOVER_SUCCESS_RATE_% BH Value")]);
                    else
                    	VFE_HSR=0.0;
                }
                ocell=orow.createCell(7);
    			ocell.setCellValue(VFE_HSR);
            
            }}
    	        br.close();
    	        FileOutputStream outExcel=new FileOutputStream(new File("C:/WORK/Final Input/Final Input/outputabc.xlsx"));           
    	        workbook.write(outExcel); //write the output data
    	        outExcel.close();   
    		    workbook.close();
    		    //=========================PIVOT MAKING AND FUNCTIONS ========================================
    		    Workbook workbook1 = null;
				try {
					workbook1 = new Workbook("C:/WORK/Final Input/Final Input/outputabc.xlsx");
				} catch (Exception e) {
					e.printStackTrace();
				}
    		    WorksheetCollection worksheets1 = workbook1.getWorksheets();
                Worksheet sheet1 = worksheets1.get("sheet1");
                String DataRange1="="+"sheet1"+"!A1:I"+(length.size()+1);
                PivotTableCollection pivotTables1 = sheet1.getPivotTables();
                int index1 = pivotTables1.add(DataRange1, "J1", "Pivot1");
                PivotTable pivotTable1 = pivotTables1.get(index1);
                pivotTable1.setRowGrand(true);
                pivotTable1.setColumnGrand(true);
                pivotTable1.setAutoFormat(true);
                pivotTable1.setAutoFormatType(PivotTableAutoFormatType.REPORT_6);
                      
                pivotTable1.addFieldToArea(PivotFieldType.ROW, 0);
           //   pivotTable.addFieldToArea(PivotFieldType.COLUMN, 0);
                pivotTable1.addFieldToArea(PivotFieldType.PAGE, 8);
           //   pivotTable1.addFieldToArea(PivotFieldType.COLUMN, 8);
                pivotTable1.addFieldToArea(PivotFieldType.DATA, 1);
                pivotTable1.addFieldToArea(PivotFieldType.DATA, 2);
                pivotTable1.addFieldToArea(PivotFieldType.DATA, 3);
                pivotTable1.addFieldToArea(PivotFieldType.DATA, 4);
                pivotTable1.addFieldToArea(PivotFieldType.DATA, 5); 
                pivotTable1.addFieldToArea(PivotFieldType.DATA, 6);
                pivotTable1.addFieldToArea(PivotFieldType.DATA, 7);
                pivotTable1.getColumnFields().add(pivotTable1.getDataField());
               // pivotTable1.getColumnFields().get(0).addCalculatedItem("Average", "=Average('ERCS_BSS_TCH_DROP BH Value','ERCS_BSS_NQI_TCH_Assign_Suc_Rate% BH Value')");
               pivotTable1.getDataFields().get(0).setFunction(ConsolidationFunction.AVERAGE);
                pivotTable1.getDataFields().get(1).setFunction(ConsolidationFunction.AVERAGE);
                pivotTable1.getDataFields().get(2).setFunction(ConsolidationFunction.AVERAGE);
               pivotTable1.getDataFields().get(3).setFunction(ConsolidationFunction.AVERAGE);
                pivotTable1.getDataFields().get(4).setFunction(ConsolidationFunction.AVERAGE);
                pivotTable1.getDataFields().get(5).setFunction(ConsolidationFunction.AVERAGE);
                pivotTable1.getDataFields().get(6).setFunction(ConsolidationFunction.AVERAGE);
                workbook1.save("C:/WORK/Final Input/Final Input/outputabc.xlsx");
                
	}
        catch (Exception e)
		{
			e.printStackTrace();
		}	
    	    
	}
	// String[] siteid = currentLine.split(",");
   
	public boolean addcolumn() throws Exception 
	{
		ArrData_Fact=new String[8];
	
		
		/*ArrData_Fact=new String[8];
        ArrData_Fact[0]= TCH_BL;
        ArrData_Fact[1]="TCH_Drop";
        ArrData_Fact[2]="Site_ID";
        ArrData_Fact[3]="SD_Bl";
        ArrData_Fact[4]="SD_drop";
        ArrData_Fact[5]="SD_Assig";
        ArrData_Fact[6]="VFE_TASR";
        ArrData_Fact[7]="VFE_HSR";
        int row=2;
        HashMap_Fact.put("Site_ID", ArrData_Fact);*/
        
    	return false;
    	
	}

	/*HashMap_Fact=new HashMap<String , String[]>();
	XSSFWorkbook workbook_input = new XSSFWorkbook(new FileInputStream(new File(PATH +"UPW 2G Worst Cell Tracker- Jul'16.xlsx")));
	XSSFSheet sheet = workbook_input.getSheet("sheet1");*/

	public void processOneSheet(String filename,String sheetname) throws Exception
	{

		 StylesTable styles = null;
		 OPCPackage pkg = OPCPackage.open(filename);
		 XSSFReader r = new XSSFReader(pkg);
		 styles = r.getStylesTable();
		 SharedStringsTable sst = r.getSharedStringsTable();
		 XMLReader parser = fetchSheetParser(sst,styles);
		 XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator)r.getSheetsData();

		while (iter.hasNext())
		{
			InputStream sheet = iter.next();
			if(sheetname.equalsIgnoreCase(iter.getSheetName()))
			{
				InputSource sheetSource = new InputSource(sheet);
				parser.parse(sheetSource);
			}
			sheet.close();
		}

		
	}

    public XMLReader fetchSheetParser(SharedStringsTable sst, StylesTable styles) throws SAXException {
        XMLReader parser =
            XMLReaderFactory.createXMLReader(
                    "org.apache.xerces.parsers.SAXParser"
            );
        ContentHandler handler = new SheetHandler(sst, styles);
        parser.setContentHandler(handler);
        return parser;
    }
	private static class SheetHandler extends DefaultHandler {
		private enum xssfDataType 
	    {  
	        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER 
	    }
		private SharedStringsTable sst;
		private String lastContents;
		private boolean nextIsString;
		private String cellRef = null;
		private int cellIndex =0;
		private int iRow = 0;
		private int iCol =0;
		
		private static int flag =0;
		private StylesTable stylesTable;
        private xssfDataType nextDataType;
        private static HashMap<Integer, String> IMAP;
    	private static String[] ROWARR;
    	private static ArrayList<String> columns ;
        private short formatIndex;
        private String formatString;
        private final DataFormatter formatter1;
        private static HashMap<String, Integer> INDEX = new HashMap<String, Integer>();
        
        
        private SheetHandler(SharedStringsTable sst, StylesTable styles) {
            this.sst = sst;     
            this.stylesTable = styles;
       //     list.add(new ArrayList<String>());
            this.nextDataType = xssfDataType.NUMBER;  
            this.formatter1 = new DataFormatter(); 
       }
	
		public void startElement(String uri, String localName, String name,
				Attributes attributes) throws SAXException 
		{
			if(name.equals("c")) {

				String cellType = attributes.getValue("t");
				if(cellType != null && cellType.equals("s")) {
					nextIsString = true;
				} 
				else 
				{
					nextIsString = false;
				}
				cellRef = attributes.getValue("r");
				int firstDigit = -1;
				for (int c = 0; c < cellRef.length(); ++c) 
				{ 
					if (Character.isDigit(cellRef.charAt(c))) 
					{ 
						firstDigit = c;
						break;
					}
				}
				cellIndex = nameToColumn(cellRef.substring(0, firstDigit)); 
				 this.nextDataType = xssfDataType.NUMBER;  
                 this.formatIndex = -1;  
                 this.formatString = null;  
                 String cellStyleStr = attributes.getValue("s");  
                 if ("b".equals(cellType))  
                     nextDataType = xssfDataType.BOOL;  
                 else if ("e".equals(cellType))  
                     nextDataType = xssfDataType.ERROR;  
                 else if ("inlineStr".equals(cellType))  
                     nextDataType = xssfDataType.INLINESTR;  
                 else if ("s".equals(cellType))  
                     nextDataType = xssfDataType.SSTINDEX;  
                 
                 else if (cellStyleStr != null) 
                 {  
                     // It's a number, but almost certainly one  
                     // with a special style or format  
                     int styleIndex = Integer.parseInt(cellStyleStr);  
                     XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);  
                     this.formatIndex = style.getDataFormat();  
                     this.formatString = style.getDataFormatString();  
                     if (this.formatString == null)  
                         this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);  
                 } 
             }
             else if(name.equals("f"))
             {
                  nextDataType = xssfDataType.FORMULA;  
             }
			
			// Clear contents cache
			lastContents = "";
	 }
		
			public void endElement(String uri, String localName, String name)throws SAXException 
			{
				
			if(nextIsString) 
			{
						int idx = Integer.parseInt(lastContents);
						lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString().trim();
						nextIsString = false;
			}
			if(name.equals("v")) 
		    {
				if(iRow == 1) 
				{
						
					
					if(lastContents.equals("#Day")){
	
						Ar.add(iCol-1);
					}
						
						
				
					iCol++;
				}
				else if(iRow>1){
					
					if(iCol==0)
					{
						
						cell = lastContents;
					}
					iCol++;
				}
				
			
              

			}
		 
		  else if(name.equals("row"))
		  {
			  if(iRow > 1)
			  {
				  try {
					  this.values();
					
				} catch (ParseException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			  }
		
			iRow++;
			iCol=0;
		 //   ROWARR = new String[columns.size()];
          }
			
			
			
			
		  }
			public void characters(char[] ch, int start, int length)throws SAXException 
			{
				lastContents += new String(ch, start, length);
			}
			private int nameToColumn(String name)
			{
				int column = -1;
				for (int i = 0; i < name.length(); ++i) 
				{
					int c = name.charAt(i);
					column = (column + 1) * 26 + c - 'A';
	        	}
				return column; 
			}
			
			void values() throws ParseException
			{
				String[] list;
				int col1 = Ar.get(0);
				int col2 = Ar.get(1);
				int col3 = Ar.get(2);
				int col4 = Ar.get(3);
				int col5 = Ar.get(4);
					
				list = HashMap_Fact.get(cell);
					
				sheet.getRow(iRow).createCell(col1).setCellValue(list[5]);
				sheet.getRow(iRow).createCell(col2).setCellValue(list[4]);
				sheet.getRow(iRow).createCell(col3).setCellValue(list[6]);
				sheet.getRow(iRow).createCell(col4).setCellValue(list[1]);
				sheet.getRow(iRow).createCell(col5).setCellValue(list[7]);
			
				}
				
			}
			
	}		
