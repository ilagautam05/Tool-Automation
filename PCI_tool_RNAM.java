package sadagi.validation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;



public class PCI_tool_RNAM 
{

	private enum xssfDataType 
    {  
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER 
    }
	private String PATH ="C:/Users/emanbaj/PCI_RNAM/";
	private static HashMap<String, Integer> INDEX ;
	private static HashMap<String, String> map_enodeB;
	private static HashMap<String, String> map_enodeB_tac;
	private static HashMap<String, String> map_enodeB_loc;
	private static HashMap<String, String> map_rbs;
	private static String[] array ={"node_name","IPV6_SIAD_BEARER_IP_DEF_ROUTER","IPV6_ENODEB_BEARER_IP","bearer_enodeb_sb_vlan_id","IPV6_ENODEB_OAM_IP",
	"oam_enodeb_siad_oam_vlan","IPV6_SIAD_OAM_IP_DEF_ROUTER","IPV6_ENODEB_OAM_IP","IPV6_VLAN_A_SUBNET_MASK","IPV6_ENODEB_SIAD_BEARER_SUB_64","bearer_siad_ip_def_router",
	"bearer_enode_b_bearer_ip","oam_enode_b_oam_ip","oam_siad_ip_def_router","vlan_a_subnet_mask"};
	private static ArrayList<ArrayList<String>> list;
	private static ArrayList<ArrayList<String>> list_market;
	private static ArrayList<ArrayList<String>> list_pci;
	private static ArrayList<ArrayList<String>> list_eutran;
	private static ArrayList<ArrayList<String>> list_losses;
	private static ArrayList<ArrayList<String>> list_edp;
	private static HashMap<String,TreeMap<Integer,ArrayList<Integer>>> map_pci ;
	private static TreeMap<Integer,ArrayList<Integer>> tree_pci ;
	private static ArrayList<Integer> list_tree;
	
	private static HashMap<Integer, String> IMAP ;
	private static String[] ROWARR;
	private static ArrayList<String> columns ;
	private static int flag =0;
	private static int count =0;
	private static String cell=null;
	private static int carrier;
	private static int sector;
	public boolean doprocess() throws Exception 
	{
		 flag =0;
		 count =0;
		 map_pci = new HashMap<String,TreeMap<Integer,ArrayList<Integer>>>();
		 tree_pci = new TreeMap<Integer,ArrayList<Integer>>();
		 list_tree = new ArrayList<Integer>();
		map_enodeB = new HashMap<String, String>();
		map_enodeB_tac = new HashMap<String, String>();
		map_enodeB_loc =new HashMap<String, String>();
		map_rbs = new HashMap<String, String>();
		list_market = new ArrayList<ArrayList<String>>();
		list = new  ArrayList<ArrayList<String>>();
		list_pci= new  ArrayList<ArrayList<String>>();
		list_eutran= new  ArrayList<ArrayList<String>>();
		list_losses=new  ArrayList<ArrayList<String>>();
		list_edp = new  ArrayList<ArrayList<String>>();
		columns = new ArrayList<String>() ;
		INDEX = new HashMap<String, Integer>();
		IMAP = new HashMap<Integer, String>();
		flag =1;
	
		this.processOneSheet(PATH +"Market Data.xlsx", "ATT National NSB Only");
		list_market = list;
		columns = new ArrayList<String>() ;
		INDEX = new HashMap<String, Integer>();
		IMAP = new HashMap<Integer, String>();
		flag =2;
		count =0;
		list = new  ArrayList<ArrayList<String>>();
		this.processOneSheet(PATH +"HOUSTON_LTE_RNDCIQ_Rev368.1(06.28.2016).xlsx", "eNB Info");
		columns = new ArrayList<String>() ;
		INDEX = new HashMap<String, Integer>();
		IMAP = new HashMap<Integer, String>();
		flag =3;
		count =0;
		list = new  ArrayList<ArrayList<String>>();
		this.processOneSheet(PATH +"HOUSTON_LTE_RNDCIQ_Rev368.1(06.28.2016).xlsx", "eUtran Parameters");
		list_eutran = list;
		columns = new ArrayList<String>() ;
		INDEX = new HashMap<String, Integer>();
		IMAP = new HashMap<Integer, String>();
		flag =4;
		count =0;
		list = new  ArrayList<ArrayList<String>>();
		this.processOneSheet(PATH +"HOUSTON_LTE_RNDCIQ_Rev368.1(06.28.2016).xlsx", "PCI");
		list_pci = list;
		columns = new ArrayList<String>() ;
		INDEX = new HashMap<String, Integer>();
		IMAP = new HashMap<Integer, String>();
		flag =5;
		count =0;
		list = new  ArrayList<ArrayList<String>>();
		this.processOneSheet(PATH +"HOUSTON_LTE_RNDCIQ_Rev368.1(06.28.2016).xlsx", "Losses and Delays");
		list_losses = list;
		columns = new ArrayList<String>() ;
		INDEX = new HashMap<String, Integer>();
		IMAP = new HashMap<Integer, String>();
		flag =6;
		count =0;
		list = new  ArrayList<ArrayList<String>>();
		for(int k=0;k<array.length;k++)
		{
		
			columns.add(array[k]);
		
		}
		
		this.processOneSheet(PATH +"EDP.xlsx", "raptor");
		list_edp = list;
		this.output();

		return false;

	}
	
	
	
	
	
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
		private SharedStringsTable sst;
		private String lastContents;
		private boolean nextIsString;
		private String cellRef = null;
		private int cellIndex =0;
		private int iRow = 0;
		private int iCol =0;
		private StylesTable stylesTable;
        private xssfDataType nextDataType;
        private short formatIndex;
        private String formatString;
        private final DataFormatter formatter1;
		
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
				if(iRow == 0) 
				{
						
					
					if(flag ==1 && (iCol==1 || iCol==3 || iCol==5 || (iCol>=9 && iCol<=18)))
					{
						columns.add(lastContents);
						INDEX.put(lastContents, cellIndex);
						IMAP.put(cellIndex, lastContents);
					}
					else if(flag ==2 && (iCol<=2 || iCol==4 || iCol==11) ){
						columns.add(lastContents);
						INDEX.put(lastContents, cellIndex);
						IMAP.put(cellIndex, lastContents);
						
					}
					else if(flag ==4 && ((iCol>=0 && iCol <=4) || iCol==6 || iCol==7) ){
						columns.add(lastContents);
						INDEX.put(lastContents, cellIndex);
						IMAP.put(cellIndex, lastContents);
						
					}
					else if(flag ==3 && ((iCol==1 || iCol==2|| iCol==4|| iCol==6 || iCol==18 ||  iCol==19 || iCol==15 || iCol==16||(iCol>=23 && iCol<=28) || iCol==33) )){
						columns.add(lastContents);
						INDEX.put(lastContents, cellIndex);
						IMAP.put(cellIndex, lastContents);
						
					}
					else if(flag ==5 && (iCol==0 || iCol>=52)){
						columns.add(lastContents);
						INDEX.put(lastContents, cellIndex);
						IMAP.put(cellIndex, lastContents);
						
					}
					else if(flag==6 &&  columns.contains(lastContents))
					{
						INDEX.put(lastContents, cellIndex);
						IMAP.put(cellIndex, lastContents);
					}
						
						
						
				
					iCol++;
				}
				
				else if(iRow > 0)
                {
                    if(INDEX.containsValue(cellIndex))
                    { 
                        switch (nextDataType)
                        {  
                        
                        case BOOL:  
                            if(!lastContents.equals(""))
                              {
                                ROWARR[columns.indexOf(IMAP.get(cellIndex))]= lastContents; 
                              }
                            break;  
          
                        case FORMULA:  
                            if(!lastContents.equals(""))
                              {
                                ROWARR[columns.indexOf(IMAP.get(cellIndex))] = lastContents;  
                              }
                            break;  
          
                        case INLINESTR:   
                            if(!lastContents.equals(""))
                                {
                                  ROWARR[columns.indexOf(IMAP.get(cellIndex))] = lastContents; 
                                }
                            break;  
          
                        case SSTINDEX:  
                            if(!lastContents.equals(""))
                            {
                                ROWARR[columns.indexOf(IMAP.get(cellIndex))] =lastContents;  
                            }
                            break;  
          
                        case NUMBER:  
                            if(!lastContents.equals(""))
                            {  
                                if (this.formatString == null || this.formatString.equals("0.00") || this.formatString.equalsIgnoreCase("General"))  
                                 {
                                    ROWARR[columns.indexOf(IMAP.get(cellIndex))] = lastContents;  
                                 }
                                else  
                                 {
                                    formatter1.setDefaultNumberFormat(new SimpleDateFormat(formatString));
                                    lastContents = formatter1.formatRawCellContents(Double.parseDouble(lastContents), this.formatIndex, this.formatString);  
                                    ROWARR[columns.indexOf(IMAP.get(cellIndex))] =lastContents;   
                                 }
                            }
                            break;  
          
                        default:  
                            ROWARR[columns.indexOf(IMAP.get(cellIndex))] =lastContents;  
                            break;  
                        } 
                    }
                }

			}
		 
		  else if(name.equals("row"))
		  {
			  if(iRow > 0)
			  {
				  try {
					  this.values(ROWARR);
					
				} catch (ParseException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			  }
		
			iRow++;
		    ROWARR = new String[columns.size()];
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
			
			void values(String[] ARR) throws ParseException
			{
				list.add(new ArrayList<String>());
			
				for(int i=0;i<columns.size();i++)
				{
					
					
					if(flag ==2){
						
						map_enodeB.put( ARR[columns.indexOf(columns.get(1))],ARR[columns.indexOf(columns.get(0))]);
						map_enodeB_tac.put( ARR[columns.indexOf(columns.get(1))],ARR[columns.indexOf(columns.get(4))]);
						map_enodeB_loc.put( ARR[columns.indexOf(columns.get(1))],ARR[columns.indexOf(columns.get(2))]);
						map_rbs.put( ARR[columns.indexOf(columns.get(1))], ARR[columns.indexOf(columns.get(3))]);
					}
					
					else if (ARR[columns.indexOf(columns.get(i))]!=null)
					{
						list.get(count).add(ARR[columns.indexOf(columns.get(i))]);
			
					}
					

					
				}
				if(flag ==4){
					
					cell =ARR[columns.indexOf(columns.get(0))].split("_")[0];
					if(ARR[columns.indexOf(columns.get(6))]!=null)
					{
					carrier= Integer.parseInt(ARR[columns.indexOf(columns.get(6))]);
					sector = Integer.parseInt(ARR[columns.indexOf(columns.get(1))]);
					
					
					if(map_pci.containsKey(cell)){
						
					tree_pci = map_pci.get(cell);
					list_tree = new ArrayList<Integer>();
					
					
					if(tree_pci.containsKey(carrier))
					{
						list_tree = tree_pci.get(carrier);
						list_tree.add(sector);
						tree_pci.put(carrier, list_tree);
						map_pci.put(cell, tree_pci);
						
					}
					else{
						list_tree = new ArrayList<Integer>();
						list_tree.add(sector);
						tree_pci.put(carrier, list_tree);
						map_pci.put(cell, tree_pci);
						
						
					}
					
					}
					else{
						
						tree_pci = new TreeMap<Integer,ArrayList<Integer>>();


							list_tree = new ArrayList<Integer>();
							list_tree.add(sector);
							tree_pci.put(carrier, list_tree);
							map_pci.put(cell, tree_pci);
							
					
					}
					}
					
					
				}
				count++;
			}
			
	}
	
	
	
	
	
	private void output() throws IOException{
		XSSFWorkbook workbook_output = new XSSFWorkbook();
		XSSFWorkbook workbook_input = new XSSFWorkbook(new FileInputStream(new File(PATH +"Input.xlsx")));
		XSSFSheet sheet_input = workbook_input.getSheet("Sheet1");
		int row = 4;
		int sheet_count =0;
		int sheet_count1 =0;
		HashMap<String,String> sheetname = new 	HashMap<String,String>();
		HashMap<String,String> sheetname_1 = new 	HashMap<String,String>();
		XSSFCellStyle greenStyle = (XSSFCellStyle) workbook_output.createCellStyle();
        byte[] rgb = new byte[3];
        rgb[0] = (byte) 0; // red
        rgb[1] = (byte) 204; // green
        rgb[2] = (byte) 255; // blue
        XSSFColor myColor = new XSSFColor(rgb);
        greenStyle.setFillForegroundColor(myColor);
        Font font = workbook_output.createFont();
        font.setBold(true);
        font.setFontName("Calibri");
    	greenStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        greenStyle.setFont(font); 
        greenStyle.setAlignment(HorizontalAlignment.CENTER);
        greenStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        HashMap<String,String> Sector_map = new HashMap<String,String>();
        HashMap<String,String> Sector_unit = new HashMap<String,String>();
        HashMap<String,String> output_power = new HashMap<String,String>();
        Font font1 = workbook_output.createFont();
    	font1.setFontName("Calibri");
    	font1.setFontHeightInPoints((short)11);
    	XSSFCellStyle style1 = (XSSFCellStyle) workbook_output.createCellStyle();
    	style1.setFont(font1);
    	style1.setWrapText(true);
		ArrayList<String> list_trav_A = new ArrayList<String>();
		ArrayList<String> list_trav_B = new ArrayList<String>();
		ArrayList<String> list_trav_C = new ArrayList<String>();
		ArrayList<String> list_trav_D = new ArrayList<String>();
		ArrayList<String> list_trav_E = new ArrayList<String>();
		ArrayList<String> list_trav_F = new ArrayList<String>();
		ArrayList<String>  cell_carriers = new ArrayList<String> ();
    	
    	
    	XSSFCellStyle orangestyle = (XSSFCellStyle) workbook_output.createCellStyle();
        byte[] rgb1 = new byte[3];
        rgb1[0] = (byte) 255; // red
        rgb1[1] = (byte) 153; // green
        rgb1[2] = (byte) 0; // blue
        XSSFColor myColor1 = new XSSFColor(rgb1);
        orangestyle.setFillForegroundColor(myColor1);
        orangestyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        orangestyle.setFont(font); 
        orangestyle.setAlignment(HorizontalAlignment.CENTER);
        orangestyle.setVerticalAlignment(VerticalAlignment.CENTER);
   
    	XSSFCellStyle greenstyle = (XSSFCellStyle) workbook_output.createCellStyle();
        byte[] rgb2 = new byte[3];
        rgb2[0] = (byte) 153; // red
        rgb2[1] = (byte) 204; // green
        rgb2[2] = (byte) 0; // blue
        XSSFColor myColor2 = new XSSFColor(rgb2);
        greenstyle.setFillForegroundColor(myColor2);
        greenstyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        
       double Decimal_lat =0.0;
       double Degrees_lat =0.0;
       double Minutes_lat =0.0;
       double Seconds_lat =0.0;
       double Milliseconds_lat =0.0;
       			
       double Decimal_long =0.0;
       double Degrees_long =0.0;
       double Minutes_long =0.0;
       double Seconds_long =0.0;
       double Milliseconds_long =0.0;

    	String rru=null;
    	String rbs = null;
    	String market;
    	String enodeB = null;
    	String ipconfig;
    	String[] columns_0 ={"dlTrafficDelay","ulTrafficDelay","dlAttenuation","ulAttenuation"};
    	
	String[] columns1 = {"Name","Node Type","Sw Release","#Template Mim Transport","eNBId","userLabel"};
		String[] columns2 ={"Name","Node Type","Sw Release","#Template ARNE","mimVersion","sitename",
"ftpAutoIntegration","ftpBackUpStore","ftpLicenseKey","ftpSwStore","ipAddress","nodeName",
"lattitude","location","longitude","siteId","worldTimeZoneId"
};
	String[] columns2_1 ={"NodeDataType(1) ManagedElement(1) mimVersion","NodeDataType(1) ManagedElement(1) sitename","NodeDataType(1) ManagedElement(1) connectivityInfo(1) ftpAutoIntegration",
				"NodeDataType(1) ManagedElement(1) connectivityInfo(1) ftpBackUpStore","NodeDataType(1)	ManagedElement(1) connectivityInfo(1) ftpLicenseKey",
				"NodeDataType(1) ManagedElement(1) connectivityInfo(1) ftpSwStore","NodeDataType(1) ManagedElement(1) connectivityInfo(1) ipAddress",
				"NodeDataType(1) ManagedElement(1) topologyInfo(1) nodeName","NodeDataType(1) siteInfo(1) lattitude","NodeDataType(1) siteInfo(1) location",
				"NodeDataType(1) siteInfo(1) longitude","NodeDataType(1) siteInfo(1) siteId","NodeDataType(1) siteInfo(1) worldTimeZoneId"};

		String[] columns3={"Name","Node Type","Sw Release","#Template Site Basic","nextHopIpAddr","ipAddress","vid","nodeLocalTimeZone"};
		
		String[] columns3_1 ={"SiteBasic(1) Ip(1) IpRoutingTable(1) StaticRoutes(1) nextHopIpAddr","SiteBasic(1) IpSystem(1) IpAccessHostEt(1) ipAddress", 
				"SiteBasic(1) IpSystem(1) Ipv6Interface(1) vid","SiteBasic(1) ManagedElementData(1) nodeLocalTimeZone"};
		
		
		String[] columns4 ={"Name","Node Type","Sw Release","#Template Site Installation","logicalName","vlanId","defaultRouter0",
				"ipAddress","networkPrefixLength"
};
		
		String[] columns4_1 ={"RbsSiteInstallationFile(1) InstallationData(1) logicalName","RbsSiteInstallationFile(1) InstallationData(1) vlanId",
				"RbsSiteInstallationFile(1) InstallationData(1) OamIpConfigurationData(1) defaultRouter0",
				"RbsSiteInstallationFile(1) InstallationData(1) OamIpConfigurationData(1) ipAddress",
				"RbsSiteInstallationFile(1) InstallationData(1) OamIpConfigurationData(1) networkPrefixLength"};
		
		String[] columns5 ={"EUtranCellFDDId","cellId","cellRange","dlChannelBandwidth","earfcndl","earfcnul","physicalLayerCellIdGroup",
				"physicalLayerSubCellId","rachRootSequence","tac","ulChannelBandwidth","userLabel","EUtranFreqRelationId","adjacentFreq","userLabel"
};
		
		Sector_unit.put("1","1");
		Sector_unit.put("2","1");
		Sector_unit.put("3","2");
		Sector_unit.put("4","3");
		Sector_unit.put("5","4");
		Sector_unit.put("6","5");
		
		XSSFSheet sheet_output = workbook_output.createSheet("Mim Transport_1");
		XSSFSheet sheet_output1 = workbook_output.createSheet("ARNE_1");
		XSSFSheet sheet_output2 = workbook_output.createSheet("Site Basic_1");
		XSSFSheet sheet_output3 = workbook_output.createSheet("Site Installation_1");
		XSSFSheet sheet_output4 =null;
		XSSFSheet sheet_output5 = null;
//		//headers
		sheet_output.createRow(0).createCell(0).setCellValue("Site Specific");
		sheet_output1.createRow(0).createCell(0).setCellValue("Site Specific");
		sheet_output2.createRow(0).createCell(0).setCellValue("Site Specific");
		sheet_output3.createRow(0).createCell(0).setCellValue("Site Specific");
		
		sheet_output.getRow(0).getCell(0).setCellStyle(orangestyle);
		sheet_output1.getRow(0).getCell(0).setCellStyle(orangestyle);
		sheet_output2.getRow(0).getCell(0).setCellStyle(orangestyle);
		sheet_output3.getRow(0).getCell(0).setCellStyle(orangestyle);

		sheet_output.createRow(3);
		sheet_output1.createRow(3);
		sheet_output2.createRow(3);
		sheet_output3.createRow(3);	
		sheet_output3.createRow(1);	
		sheet_output3.createRow(2);	
		sheet_output2.createRow(1);	
		sheet_output2.createRow(2);	
		sheet_output1.createRow(1);	
		sheet_output1.createRow(2);	
		sheet_output.createRow(1);	
		sheet_output.createRow(2);
		
	
		sheet_output.createRow(1).createCell(4).setCellValue("ManagedElement(1) ENodeBFunction(1) eNBId");
		sheet_output.getRow(1).createCell(5).setCellValue("ManagedElement(1) ENodeBFunction(1) userLabel");
		sheet_output.getRow(1).getCell(4).setCellStyle(style1);	
		sheet_output.getRow(1).getCell(5).setCellStyle(style1);	
		sheet_output.getRow(2).createCell(4).setCellStyle(greenstyle);
		sheet_output.getRow(2).createCell(5).setCellStyle(greenstyle);
		
		
		for(int j=0;j<columns2.length;j++){
			
			sheet_output1.getRow(3).createCell(j).setCellValue(columns2[j]);
			sheet_output1.getRow(3).getCell(j).setCellStyle(greenStyle);
			
			if(j>=4){
				sheet_output1.getRow(1).createCell(j).setCellValue(columns2_1[j-4]);
				sheet_output1.getRow(1).getCell(j).setCellStyle(style1);	
				sheet_output1.getRow(2).createCell(j).setCellStyle(greenstyle);	
			}
		}
	
		for(int j=0;j<columns1.length;j++){
			
			sheet_output.getRow(3).createCell(j).setCellValue(columns1[j]);
			sheet_output.getRow(3).getCell(j).setCellStyle(greenStyle);
		}
		for(int j=0;j<columns3.length;j++){
			
			sheet_output2.getRow(3).createCell(j).setCellValue(columns3[j]);
			sheet_output2.getRow(3).getCell(j).setCellStyle(greenStyle);
			if(j>=4){
				sheet_output2.getRow(1).createCell(j).setCellValue(columns3_1[j-4]);
				sheet_output2.getRow(1).getCell(j).setCellStyle(style1);	
				sheet_output2.getRow(2).createCell(j).setCellStyle(greenstyle);	
			}
		}
		for(int j=0;j<columns4.length;j++)
		{
			sheet_output3.getRow(3).createCell(j).setCellValue(columns4[j]);
			sheet_output3.getRow(3).getCell(j).setCellStyle(greenStyle);
			if(j>=4){
				sheet_output3.getRow(1).createCell(j).setCellValue(columns4_1[j-4]);
				sheet_output3.getRow(1).getCell(j).setCellStyle(style1);	
				sheet_output3.getRow(2).createCell(j).setCellStyle(greenstyle);	
		}
		
	}
		
		
		
		String config=null;
		String str1 =null;
		int row1 = 0;
		int row2 = 0;
		for(int k=1;k<sheet_input.getPhysicalNumberOfRows();k++)
		{
		list_trav_A = new ArrayList<String>();
		list_trav_B = new ArrayList<String>();
		list_trav_C = new ArrayList<String>();
		list_trav_D = new ArrayList<String>();
		list_trav_E = new ArrayList<String>();
		list_trav_F = new ArrayList<String>();
		cell_carriers = new ArrayList<String>();
		Sector_map = new HashMap<String,String>();
		sheet_output.createRow(row);
		sheet_output1.createRow(row);
		sheet_output2.createRow(row);
		sheet_output3.createRow(row);
		market = sheet_input.getRow(k).getCell(1).getStringCellValue();
		enodeB = sheet_input.getRow(k).getCell(0).getStringCellValue();
		ipconfig =sheet_input.getRow(k).getCell(2).getStringCellValue();
		//Mim Transport_1
		config= "RN_ATT_";
		str1 ="SE_ATT_";
		tree_pci = map_pci.get(enodeB);
		
		for(Map.Entry<Integer, ArrayList<Integer>> EN : tree_pci.entrySet())
		{	
			list_tree= tree_pci.get(EN.getKey());
			config=config+list_tree.toString().replace(",","").replace("]", "").replace("[", "").replace(" ", "")+"_"+EN.getKey()+"C_";
			str1=str1+list_tree.toString().replace(",","").replace("]", "").replace("[", "").replace(" ", "")+"_"+EN.getKey()+"C_";
			
		}
		sheet_output.getRow(row).createCell(0).setCellValue(enodeB);
		sheet_output.getRow(row).createCell(1).setCellValue("ENodeB");
		//ARNE_1
		sheet_output1.getRow(row).createCell(0).setCellValue(enodeB);
		sheet_output1.getRow(row).createCell(1).setCellValue("ENodeB");
		
		sheet_output2.getRow(row).createCell(0).setCellValue(enodeB);
		sheet_output2.getRow(row).createCell(1).setCellValue("ENodeB");
		
		sheet_output3.getRow(row).createCell(0).setCellValue(enodeB);
		sheet_output3.getRow(row).createCell(1).setCellValue("ENodeB");

		for(int l=0;l<list_market.size();l++)
		{
			if(list_market.get(l).get(0)!=null && market.equals(list_market.get(l).get(0)))
			{
				config = config + "V1_" + list_market.get(l).get(2);
				if(!(sheetname.containsKey(config)))
				{
					sheet_count++;
					sheetname.put(config, "Mim Radio_"+sheet_count);
					sheet_output4 = workbook_output.createSheet("Mim Radio_"+sheet_count);
					sheet_output4.createRow(0).createCell(0).setCellValue("Site Specific");
					sheet_output4.getRow(0).getCell(0).setCellStyle(orangestyle);
					sheet_output4.createRow(3);	
					sheet_output4.createRow(1);	
					sheet_output4.createRow(2);	

					sheet_output4.getRow(3).createCell(0).setCellValue("Name");
					sheet_output4.getRow(3).createCell(1).setCellValue("Node Type");
					sheet_output4.getRow(3).createCell(2).setCellValue("Sw Release");
					sheet_output4.getRow(3).createCell(3).setCellValue("#Template Mim Radio");
					row1 = 4;
					sheet_output4.createRow(row1);
					sheet_output4.getRow(row1).createCell(0).setCellValue(enodeB);
					sheet_output4.getRow(row1).createCell(1).setCellValue("ENodeB");
					
				}			
				else if(sheetname.containsKey(config))
				{
					sheet_output4 = workbook_output.getSheet(sheetname.get(config));
					row1 = sheet_output4.getPhysicalNumberOfRows();
					sheet_output4.createRow(row1);
					sheet_output4.getRow(row1).createCell(0).setCellValue(enodeB);
					sheet_output4.getRow(row1).createCell(1).setCellValue("ENodeB");
				}
				
				sheet_output.getRow(row).createCell(2).setCellValue(list_market.get(l).get(2));
				sheet_output.getRow(row).createCell(3).setCellValue(list_market.get(l).get(10));
				sheet_output.getRow(row).createCell(4).setCellValue(map_enodeB.get(enodeB));
				sheet_output.getRow(row).createCell(5).setCellValue(enodeB);
				sheet_output1.getRow(row).createCell(2).setCellValue(list_market.get(l).get(2));
				sheet_output1.getRow(row).createCell(3).setCellValue(list_market.get(l).get(9));
				sheet_output1.getRow(row).createCell(4).setCellValue(list_market.get(l).get(1));
				sheet_output1.getRow(row).createCell(5).setCellValue(enodeB);
				sheet_output1.getRow(row).createCell(6).setCellValue(list_market.get(l).get(8));
				sheet_output1.getRow(row).createCell(7).setCellValue(list_market.get(l).get(5));
				sheet_output1.getRow(row).createCell(8).setCellValue(list_market.get(l).get(7));
				sheet_output1.getRow(row).createCell(9).setCellValue(list_market.get(l).get(6));
				sheet_output1.getRow(row).createCell(11).setCellValue(enodeB);
				sheet_output1.getRow(row).createCell(13).setCellValue(map_enodeB_loc.get(enodeB));
	
				sheet_output1.getRow(row).createCell(15).setCellValue(enodeB);
				sheet_output1.getRow(row).createCell(16).setCellValue(list_market.get(l).get(3));
				sheet_output2.getRow(row).createCell(2).setCellValue(list_market.get(l).get(2));
				sheet_output2.getRow(row).createCell(3).setCellValue(list_market.get(l).get(11));
				sheet_output2.getRow(row).createCell(7).setCellValue(list_market.get(l).get(4));
				
				sheet_output3.getRow(row).createCell(2).setCellValue(list_market.get(l).get(2));
				sheet_output3.getRow(row).createCell(3).setCellValue(list_market.get(l).get(12));
				sheet_output3.getRow(row).createCell(4).setCellValue(enodeB);
				
				sheet_output4.getRow(row1).createCell(2).setCellValue(list_market.get(l).get(2));

//				sheet_output5.getRow(row2).createCell(2).setCellValue(list_market.get(l).get(2));

				int count = 0;       
				       
			for(int j=0;j<list_pci.size();j++){
				
				if(list_pci.get(j).get(0).contains(enodeB)){
					
				

					
					
					
					Sector_map.put(list_pci.get(j).get(0), list_pci.get(j).get(6));
					
					if(row1==4)
					{
					
					for(int p=0;p<columns5.length;p++)
					{
						
						sheet_output4.getRow(3).createCell(p+4 + (count*15)).setCellValue(columns5[p]);
						if(p<columns5.length-3)
						{
						
						sheet_output4.getRow(1).createCell(p+4 + (count*15)).setCellValue("ManagedElement(1) ENodeBFunction(1)"
								+"EUtranCellFDD("+(list_pci.get(j).get(6))+") "+columns5[p]);
								
						}
						
						else if(p>=columns5.length-3){
							
							sheet_output4.getRow(1).createCell(p+4 + (count*15)).setCellValue("ManagedElement(1) ENodeBFunction(1)"
									+"EUtranCellFDD("+(list_pci.get(j).get(6))+") EUtranFreqRelation(1) "+columns5[p]);
							
						}
						
					}
					}
			
					sheet_output4.getRow(row1).createCell(4 + (count*15)).setCellValue( list_pci.get(j).get(0));
					
					sheet_output4.getRow(row1).createCell(15 + (count*15)).setCellValue( list_pci.get(j).get(0));
					sheet_output4.getRow(row1).createCell(18 + (count*15)).setCellValue( list_pci.get(j).get(0));
					
					sheet_output4.getRow(row1).createCell(5 + (count*15)).setCellValue( list_pci.get(j).get(2));
					
					sheet_output4.getRow(row1).createCell(10 + (count*15)).setCellValue( list_pci.get(j).get(3));
					sheet_output4.getRow(row1).createCell(11 + (count*15)).setCellValue( list_pci.get(j).get(4));
					sheet_output4.getRow(row1).createCell(12 + (count*15)).setCellValue( list_pci.get(j).get(5));
					sheet_output4.getRow(row1).createCell(13 + (count*15)).setCellValue(map_enodeB_tac.get(enodeB));
		
					count++;
				}
				
				
				
			}
			
			count =0;
			
			
			for(int j=0;j<list_eutran.size();j++){
				
				if(list_eutran.get(j).get(0).contains(enodeB)){

					sheet_output4.getRow(row1).createCell(6 + (count*15)).setCellValue( list_eutran.get(j).get(3));
					
					sheet_output4.getRow(row1).createCell(7 + (count*15)).setCellValue( list_eutran.get(j).get(10));
					sheet_output4.getRow(row1).createCell(8 + (count*15)).setCellValue(  list_eutran.get(j).get(8));
					
					sheet_output4.getRow(row1).createCell(9 + (count*15)).setCellValue( list_eutran.get(j).get(9));
					
					sheet_output4.getRow(row1).createCell(14 + (count*15)).setCellValue( list_eutran.get(j).get(11));
					sheet_output4.getRow(row1).createCell(16 + (count*15)).setCellValue( list_eutran.get(j).get(8));
					sheet_output4.getRow(row1).createCell(17 + (count*15)).setCellValue( list_eutran.get(j).get(8));
					
				   Decimal_lat =  Double.parseDouble(list_eutran.get(j).get(1));
				   Degrees_lat =   Double.parseDouble(new DecimalFormat("#.##").format(Decimal_lat));
				   //=((((E9*3600)+(F9*60)+G9)*1000)+H9)
				   Minutes_lat =Double.parseDouble(new DecimalFormat("#.##").format((Decimal_lat - Degrees_lat)*60));
				   Seconds_lat = Double.parseDouble(new DecimalFormat("#.##").format((Decimal_lat - Degrees_lat)*3600 - (Minutes_lat*60)));
				   Milliseconds_lat = ((Decimal_lat - Degrees_lat)*3600 - (Minutes_lat*60)) - Seconds_lat ;
				       			
				   Decimal_long =  Double.parseDouble(list_eutran.get(j).get(2));
				   Degrees_long =   Double.parseDouble(new DecimalFormat("#.##").format(Decimal_long));
				   Minutes_long =Double.parseDouble(new DecimalFormat("#.##").format((Decimal_long - Degrees_long)*60));
				   Seconds_long = Double.parseDouble(new DecimalFormat("#.##").format((Decimal_long - Degrees_long)*3600 - (Minutes_long*60)));
				   Milliseconds_long = ((Decimal_long - Degrees_long)*3600 - (Minutes_long*60)) - Seconds_long ;
					
					
					sheet_output1.getRow(row1).createCell(12).setCellValue(((((Degrees_lat*3600)+(Minutes_lat*60)+Seconds_lat)*1000)+Milliseconds_lat));
					sheet_output1.getRow(row1).createCell(14).setCellValue(((((Degrees_long*3600)+(Minutes_long*60)+Seconds_long)*1000)+Milliseconds_long));
					
					 rbs = list_eutran.get(j).get(6).split("/")[0];
					 if(list_eutran.get(j).get(14).contains("RRUS11") && ( rru==null || !(rru.contains("RRU")))){
						 if(rru==null)
						 rru = "RRU";
						 else{
							 rru = rru + "RRU";
							 
							 
						 }
					 }
					 else if(list_eutran.get(j).get(14).contains("RRUSA2")){
						 
						 rru = rru+"RRUA2";
					 }
					count++;
					
				}
				
				
				
			}
			str1 = str1 + map_rbs.get(enodeB)+"_"+rru+"_"+rbs+"_V1_" + list_market.get(l).get(2);
			
			if(!(sheetname_1.containsKey(str1)))
			{
				sheet_count1++;
				sheetname_1.put(str1, "Site Equipment_"+sheet_count1);
				sheet_output5 = workbook_output.createSheet("Site Equipment_"+sheet_count1);
				sheet_output5.createRow(0).createCell(0).setCellValue("Site Specific");
				sheet_output5.getRow(0).getCell(0).setCellStyle(orangestyle);
				sheet_output5.createRow(3);	
				sheet_output5.createRow(1);	
				sheet_output5.createRow(2);	
				sheet_output5.getRow(3).createCell(0).setCellValue("Name");
				sheet_output5.getRow(3).createCell(1).setCellValue("Node Type");
				sheet_output5.getRow(3).createCell(2).setCellValue("Sw Release");
				sheet_output5.getRow(3).createCell(3).setCellValue("#Template Site Equipment");
				row2 = 4;
				sheet_output5.createRow(row2);
				sheet_output5.getRow(row2).createCell(0).setCellValue(enodeB);
				sheet_output5.getRow(row2).createCell(1).setCellValue("ENodeB");
				
			}			
			else if(sheetname_1.containsKey(str1))
			{
				sheet_output5 = workbook_output.getSheet(sheetname_1.get(str1));
				row2 = sheet_output5.getPhysicalNumberOfRows();
				sheet_output5.createRow(row2);
				sheet_output5.getRow(row2).createCell(0).setCellValue(enodeB);
				sheet_output5.getRow(row2).createCell(1).setCellValue("ENodeB");
			}
			
			
			count--;
			
			int cell = 19 + (count*15);
			count=0;
			for(int j=0;j<list_eutran.size();j++){
				
				if(list_eutran.get(j).get(0).contains(enodeB)){

					
					if(!(cell_carriers.contains(Sector_unit.get(Sector_map.get(list_eutran.get(j).get(0)))))){
		
						
						output_power.put(Sector_unit.get(Sector_map.get(list_eutran.get(j).get(0))),   list_eutran.get(j).get(12));

						if(row1==4)
						{
					sheet_output4.getRow(3).createCell(cell).setCellValue("noOfRxAntennas");
		
					sheet_output4.getRow(1).createCell(cell).setCellValue("ManagedElement(1) ENodeBFunction(1)" + "SectorCarrier("+Sector_map.get(list_eutran.get(j).get(0))+")  " +"noOfRxAntennas");
					sheet_output4.getRow(row1).createCell(cell).setCellValue( list_eutran.get(j).get(4));
					
					sheet_output4.getRow(3).createCell(cell+1).setCellValue("noOfTxAntennas");
					sheet_output4.getRow(3).createCell(cell+2).setCellValue("partOfSectorPower");
					
					sheet_output4.getRow(1).createCell(cell+1).setCellValue("ManagedElement(1) ENodeBFunction(1)" + "SectorCarrier("+Sector_map.get(list_eutran.get(j).get(0))+")   noOfTxAntennas");
					sheet_output4.getRow(1).createCell(cell+2).setCellValue("ManagedElement(1) ENodeBFunction(1)" + "SectorCarrier("+Sector_map.get(list_eutran.get(j).get(0))+")  partOfSectorPower");
					
						}
					sheet_output4.getRow(row1).createCell(cell+1).setCellValue( list_eutran.get(j).get(5));
					sheet_output4.getRow(row1).createCell(cell+2).setCellValue(  list_eutran.get(j).get(13));

					
					
					
					if(count == 0 && (Sector_map.get(list_eutran.get(j).get(0)).equals("1") || Sector_map.get(list_eutran.get(j).get(0)).equals("2")))
					{
						if(row2==4)
						{
						sheet_output5.getRow(3).createCell(4).setCellValue("mechanicalAntennaTilt");
						sheet_output5.getRow(1).createCell(4).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) AntennaUnit(1) mechanicalAntennaTilt");
						}
						sheet_output5.getRow(row2).createCell(4).setCellValue( list_eutran.get(j).get(7));
						count++;
					}
					else{
						if(row2==4)
						{
						sheet_output5.getRow(3).createCell(4+count).setCellValue("mechanicalAntennaTilt");
						sheet_output5.getRow(1).createCell(4+count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) AntennaUnit("+Sector_unit.get(Sector_map.get(list_eutran.get(j).get(0)))+") mechanicalAntennaTilt");
						}
						sheet_output5.getRow(row2).createCell(4+count).setCellValue( list_eutran.get(j).get(7));
						count++;
					
					
					
					}
						cell_carriers.add(Sector_unit.get(Sector_map.get(list_eutran.get(j).get(0))));
					
					cell = cell+3;
					
				}
					
				}
				
				
				
			}
			
		




			if(row1==4)
			{
		
			for(int j=0;j<sheet_output4.getRow(3).getPhysicalNumberOfCells();j++)
			{

				sheet_output4.getRow(3).getCell(j).setCellStyle(greenStyle);
				if(j>=4){
					
					sheet_output4.getRow(1).getCell(j).setCellStyle(style1);	
					sheet_output4.getRow(2).createCell(j).setCellStyle(greenstyle);	
				}
			}
			}
	
	
 			
			cell_carriers = new ArrayList<String>();
		
			for(int j=0;j<list_losses.size();j++)
			{
				try{
			
				if(list_losses.get(j).get(0).contains(enodeB))
				{
					
					
					if(list_losses.get(j).get(0).contains("A"))
						
						list_trav_A.add(Sector_map.get(list_losses.get(j).get(0))+"-"+list_losses.get(j).get(0)+"-"+j);
						
					
					else if(list_losses.get(j).get(0).contains("B"))
					
						
						list_trav_B.add(Sector_map.get(list_losses.get(j).get(0))+"-"+list_losses.get(j).get(0)+"-"+j);
						
					
					else if(list_losses.get(j).get(0).contains("C"))
					
						
						list_trav_C.add(Sector_map.get(list_losses.get(j).get(0))+"-"+list_losses.get(j).get(0)+"-"+j);
						
					if(list_losses.get(j).get(0).contains("D"))
						
						list_trav_D.add(Sector_map.get(list_losses.get(j).get(0))+"-"+list_losses.get(j).get(0)+"-"+j);
						
					
					else if(list_losses.get(j).get(0).contains("E"))
					
						
						list_trav_E.add(Sector_map.get(list_losses.get(j).get(0))+"-"+list_losses.get(j).get(0)+"-"+j);
						
					
					else if(list_losses.get(j).get(0).contains("F"))
					
						
						list_trav_F.add(Sector_map.get(list_losses.get(j).get(0))+"-"+list_losses.get(j).get(0)+"-"+j);
					
					
					if(!(cell_carriers.contains(Sector_map.get(list_losses.get(j).get(0)))))
					cell_carriers.add(Sector_map.get(list_losses.get(j).get(0)));
					
				}
				}
				catch(Exception e){
					
					System.out.print(j);
				}
			}
			
			
			Collections.sort(list_trav_A);
			Collections.sort(list_trav_B);
			Collections.sort(list_trav_C);
			Collections.sort(list_trav_D);
			Collections.sort(list_trav_E);
			Collections.sort(list_trav_F);
			Collections.sort(cell_carriers);
			String losses=null;
			count = count + 4;
			int rfbranch =1;
			
			
			for(String cells: cell_carriers){
				
				
				if(list_trav_A!=null && list_trav_A.size()>0 && list_trav_A.get(0).startsWith(cells))
				{
				rfbranch =1;
					
					for(String A: list_trav_A){
						
						if(A.startsWith("1") || A.startsWith("2"))
						{
							
							for(int h=0;h<2;h++)
							{
							
							for(int o=0;o<columns_0.length;o++)
							{
								if(row2==4)
								{
							sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
							sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);
								}
							for(int i=0;i<15;i++)
							{
								
								if(losses==null)
								{
								losses = 	list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
								
								else{
									
									losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
							}
							
							sheet_output5.getRow(row2).createCell(count).setCellValue(losses);	
							losses = null;
							count++;
							
							
							}
							rfbranch++;
							}
							
						}
						else{
						
							
							for(int h=0;h<4;h++)
						{
							
							for(int o=0;o<columns_0.length;o++)
							{
								if(row2==4)
								{
							sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
							sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);
								}
							
						for(int i=0;i<15;i++)
							{
								
								if(losses==null)
								{
								losses = 	list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
								
								else{
									
								losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
							}
							}
							sheet_output5.getRow(row2).createCell(count).setCellValue(losses);		
							losses=null;
							count++;
						
						
						}
						rfbranch++;
							}
							
							
							
					}
						
						
					}
					
					
					
					
					list_trav_A = null;
					losses = null;
				
				
				}
			
			if(list_trav_B!=null && list_trav_B.size()>0 && list_trav_B.get(0).startsWith(cells))
				{
					rfbranch =1;
					for(String A: list_trav_B){
					
						if(A.startsWith("1") || A.startsWith("2"))
						{
							
							for(int h=0;h<2;h++)
							{
							
							for(int o=0;o<columns_0.length;o++)
							{
								if(row2==4)
								{
							sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
							sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);
								}
							for(int i=0;i<15;i++)
							{
								
								if(losses==null)
								{
							losses = 	list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
								
								else{
									
									losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
							}
							
							sheet_output5.getRow(row2).createCell(count).setCellValue(losses);	
							losses = null;
							count++;
							
							
							}
							rfbranch++;
							}
							
						}
						else{
						
						
							for(int h=0;h<4;h++)
							{
							
							for(int o=0;o<columns_0.length;o++)
							{		
								if(row2==4)
								{
							sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
							sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);							
								}
							for(int i=0;i<15;i++)
							{
								
								if(losses==null)
								{
									try{
								losses = 	list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;}
									catch(Exception e){
										
										System.out.print(enodeB);
									}
								}
								
								else{
									
									losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
							}
							
							sheet_output5.getRow(row2).createCell(count).setCellValue(losses);		
							losses = null;
							count++;
							
							
							}
							rfbranch++;
							}
							
							
							
						}
						
						
					}
					
					list_trav_B = null;

					losses =null;
					
				}
				if(list_trav_C!=null && list_trav_C.size()>0 && list_trav_C.get(0).startsWith(cells))
				{
					rfbranch =1;
				
					for(String A: list_trav_C){
				
					if(A.startsWith("1") || A.startsWith("2"))
						{
							
							for(int h=0;h<2;h++)
							{
						
						for(int o=0;o<columns_0.length;o++)
						{
							if(row2==4)
							{
							sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
							sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);
							}
							for(int i=0;i<15;i++)
							{
								
								if(losses==null)
							{
								losses = 	list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
							}
								
								else{
									
									losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
						}
							
							sheet_output5.getRow(row2).createCell(count).setCellValue(losses);		
						losses = null;
							count++;
													
							}
							rfbranch++;
						}
						
						}
						else{
							
							for(int h=0;h<4;h++)
							{						
							for(int o=0;o<columns_0.length;o++)
							{		
								if(row2==4)
								{
								sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
								sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);					
								}
							for(int i=0;i<15;i++)
							{
							
								if(losses==null)
								{
								losses = 	list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
								
							else{
									
									losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
							}
						
						sheet_output5.getRow(row2).createCell(count).setCellValue(losses);		
							losses = null;
							count++;
							
							
						}
							rfbranch++;
							}
							
						}
						
						
					}
					
					
					
					list_trav_C = null;

					losses =null;
				
				}
				if(list_trav_D!=null && list_trav_D.size()>0 && list_trav_D.get(0).startsWith(cells))
				{
					rfbranch =1;
					
					for(String A: list_trav_D){
					
					if(A.startsWith("1") || A.startsWith("2"))
						{
							
							for(int h=0;h<2;h++)
							{
							
						for(int o=0;o<columns_0.length;o++)
						{
							if(row2==4)
							{
						sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
							sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);
							}
							for(int i=0;i<15;i++)
							{
								
								if(losses==null)
							{
								losses = 	list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
							}
								
								else{
									
									losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
						}
							
							sheet_output5.getRow(row2).createCell(count).setCellValue(losses);		
						losses = null;
							count++;
							
							
							}
							rfbranch++;
						}
							
						}
						else{
							
							
							for(int h=0;h<4;h++)
							{
							
							for(int o=0;o<columns_0.length;o++)
							{
								if(row2==4)
								{
						sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
							sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);
								}
							for(int i=0;i<15;i++)
							{
							
								if(losses==null)
								{
								losses = 	list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
								
							else{
									
									losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
							}
						
						sheet_output5.getRow(row2).createCell(count).setCellValue(losses);		
							losses = null;
							count++;
							
						
						}
							rfbranch++;
							}
						
						}
						
						
					}
					
					
					
					list_trav_D = null;

					losses =null;
				
				}
				if(list_trav_E!=null && list_trav_E.size()>0 && list_trav_E.get(0).startsWith(cells))
				{
					rfbranch =1;
					for(String A: list_trav_E){
					
						if(A.startsWith("1") || A.startsWith("2"))
						{
							
							for(int h=0;h<2;h++)
							{
							
							for(int o=0;o<columns_0.length;o++)
							{
								if(row2==4)
								{
							sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
							sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);
								}
							for(int i=0;i<15;i++)
							{
								
								if(losses==null)
								{
							losses = list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								
								}
								
								else{
									
							losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								
								}
							}
							
							sheet_output5.getRow(row2).createCell(count).setCellValue(losses);	
							losses = null;
							count++;
							
							
							}
							rfbranch++;
							}
							
						}
						else{
						
						
							for(int h=0;h<4;h++)
							{
							
							for(int o=0;o<columns_0.length;o++)
							{	
								if(row2==4)
								{
							sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
							sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);							
								}
							for(int i=0;i<15;i++)
							{
								
								if(losses==null)
								{
								losses = 	list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
								
								else{
									
									losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
							}
							
							sheet_output5.getRow(row2).createCell(count).setCellValue(losses);		
							losses = null;
							count++;
							
							
							}
							rfbranch++;
							}
							
							
							
						}
						
						
					}
					
					list_trav_E = null;

					losses =null;
					
				}
				if(list_trav_F!=null && list_trav_F.size()>0 && list_trav_F.get(0).startsWith(cells))
				{
					rfbranch =1;
					for(String A: list_trav_F){
					
						if(A.startsWith("1") || A.startsWith("2"))
						{
							
							for(int h=0;h<2;h++)
							{
							
							for(int o=0;o<columns_0.length;o++)
							{
								if(row2==4)
								{
							sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
							sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);
								}
							for(int i=0;i<15;i++)
							{
								
								if(losses==null)
								{
							losses = 	list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
								
								else{
									
									losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
							}
							
							sheet_output5.getRow(row2).createCell(count).setCellValue(losses);	
							losses = null;
							count++;
							
							
							}
							rfbranch++;
							}
							
						}
						else{
						
						
							for(int h=0;h<4;h++)
							{
							
							for(int o=0;o<columns_0.length;o++)
							{				
								if(row2==4)
								{
							sheet_output5.getRow(3).createCell(count).setCellValue(columns_0[o]);
							sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) CommonAntennaSystem(1) AntennaUnitGroup(1) RfBranch("+rfbranch+")"+columns_0[o]);							
								}
								
								for(int i=0;i<15;i++)
							{
								
								if(losses==null)
								{
								losses = 	list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
								
								else{
									
									losses = losses +","+ list_losses.get(Integer.parseInt(A.split("-")[2])).get(o+1+(4*h)) ;
								}
							}
							
							sheet_output5.getRow(row2).createCell(count).setCellValue(losses);		
							losses = null;
							count++;
							
							
							}
							rfbranch++;
							}
							
							
							
						}
						
						
					}
					
					list_trav_F = null;

					losses =null;
					
				}
			}
		
			if(row2==4)
			{
		sheet_output5.getRow(3).createCell(count).setCellValue("site");	
		sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) NodeData(1) site");	
			
			}
			
			
		sheet_output5.getRow(row2).createCell(count).setCellValue(enodeB);		
			count++;
				
			for(Map.Entry<String, String> EN : output_power.entrySet())
			{
				if(row2==4)
				{
				sheet_output5.getRow(3).createCell(count).setCellValue("configuredOutputPower");
				sheet_output5.getRow(1).createCell(count).setCellValue("SiteEquipment(1) SectorEquipment(1) configuredOutputPower("+EN.getKey()+")");
				}
				sheet_output5.getRow(row2).createCell(count).setCellValue(EN.getValue());
				count++;
			}
				
			
				
		
				sheet_output4.getRow(row1).createCell(3).setCellValue(config);
				sheet_output5.getRow(row2).createCell(3).setCellValue(str1);
				
				for(int m=0;m<list_edp.size();m++)
				{
					
					if(list_edp.get(m).get(0).equals(enodeB))
					{
					
				if(ipconfig.equals("IPv6")){
					
					sheet_output1.getRow(row).createCell(10).setCellValue(list_edp.get(m).get(4));
					sheet_output2.getRow(row).createCell(4).setCellValue(list_edp.get(m).get(1));
					sheet_output2.getRow(row).createCell(5).setCellValue(list_edp.get(m).get(2));
					sheet_output2.getRow(row).createCell(6).setCellValue(list_edp.get(m).get(3));
					
					sheet_output3.getRow(row).createCell(5).setCellValue(list_edp.get(m).get(5));
					sheet_output3.getRow(row).createCell(6).setCellValue(list_edp.get(m).get(6));
					sheet_output3.getRow(row).createCell(7).setCellValue(list_edp.get(m).get(4));
					sheet_output3.getRow(row).createCell(8).setCellValue(list_edp.get(m).get(8).substring(1, list_edp.get(m).get(8).length()));
					}
				else if(ipconfig.equals("IPv4"))
				{
					sheet_output1.getRow(row).createCell(10).setCellValue(list_edp.get(m).get(12));
					sheet_output2.getRow(row).createCell(4).setCellValue(list_edp.get(m).get(10));
					sheet_output2.getRow(row).createCell(5).setCellValue(list_edp.get(m).get(11));
					sheet_output2.getRow(row).createCell(6).setCellValue(list_edp.get(m).get(3));
					
					sheet_output3.getRow(row).createCell(5).setCellValue(list_edp.get(m).get(5));
					sheet_output3.getRow(row).createCell(6).setCellValue(list_edp.get(m).get(13));
					sheet_output3.getRow(row).createCell(7).setCellValue(list_edp.get(m).get(12));
					sheet_output3.getRow(row).createCell(8).setCellValue(list_edp.get(m).get(14));
					
				}
				break;
				}
				
				}
				
			}
			
			
		}
		
		
	if(row2==4)
	{
	
		for(int j=0;j<sheet_output5.getRow(3).getPhysicalNumberOfCells();j++)
		{

		sheet_output5.getRow(3).getCell(j).setCellStyle(greenStyle);
		if(j>=4){
				
			sheet_output5.getRow(1).getCell(j).setCellStyle(style1);	
		sheet_output5.getRow(2).createCell(j).setCellStyle(greenstyle);	
		}
		}
	
		
	}	
		
		row++;
		}
		
		
		
		
		FileOutputStream outFile = new FileOutputStream(new File(PATH +"Output Sheet.xlsx"));
		workbook_output.write(outFile);
		outFile.close();
		workbook_input.close();
		workbook_output.close();
	}
	
	
}
