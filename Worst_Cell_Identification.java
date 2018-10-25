package sadagi.ericsson.softhuman.validation.model;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.LinkedList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Worst_Cell_Identification {
	private int Short_name_Index=0;
	private int ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_Index=0;
	private int ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_Index=0;
	private int ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_Index=0;
	private int ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_Index=0;
	private int Cell_Name_Index=0;
	private int Handover_Success_Rate_Index=0;
	private int SDCCH_Assignment_Success_Index=0;
	private int TCH_Assignment_Success_Index=0;
	private int SDCCH_Completion_Rate_Index=0;
	private int TCH_Completion_Rate_Index=0;
	private int RX_Quality_DL_Index=0;
	
	private int Cell_Name_Index1=0;
	private int SDCCH_Completion_Rate_Index1=0;
	private int TCH_Completion_Rate_Index1=0;
	private int Handover_Success_Rate_Index1=0;
	private int SDCCH_Assignment_Success_Index1=0;
	private int TCH_Assignment_Success_Index1=0;
	private int DL_QL_Index1=0;
	private int UtranCell_Index=0;
	private int DCR_PS_Index=0;
	private int DCR_CS_Index=0;
	private int RRC_Succ_CS_Rate_Index=0;
	private int RRC_Succ_PS_Rate_Index=0;
	private int RAB_Succ_CS_Rate_Index=0;
	private int RAB_Succ_PS_Rate_Index=0;
	private int CSSR_Index=0;
	private int DSSR_Index=0;
	private LinkedList<String> List_2G=null;
	private LinkedList<String> List_3G=null;
	private LinkedHashMap<String, Double> Hash_Map_2G=null;
	private LinkedHashMap<String, Double> Hash_Map_3G=null;
	
	
	public void doProcess() 
	{
		try
		{
			String path="C:\\MY Work\\myPro\\SoftHumanActivity\\WebContent\\Excel\\Worst_Cell\\";
			File f=new File(path+"Output Worst Cell.xlsx");
			if(f.exists())
			{
				this.NewCode();//TRENDING
			}
			else
			{
				this.Worst_Cell_Identification_Read();//NEW ONE
			}
			
		}
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}
	public void Worst_Cell_Identification_Read() 
	{
		String path="C:\\MY Work\\myPro\\SoftHumanActivity\\WebContent\\Excel\\Worst_Cell\\";
		File DBBH_2G=null;
		File BBH_2G=null;
		File NBH_2G=null;
		File NBH_3G=null;
		System.out.println("Not Ok");
		try
		{	
			DBBH_2G=new File(path+"ERCS_BSS_CELL_WISE_CELL_DBBH_UPW_09122015.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(DBBH_2G));
			XSSFSheet DBBH_2G_Sheet = wb.getSheetAt(0);
			
			BBH_2G=new File(path+"NQI_BBH_REPORT_UPW_new.xlsx");
			XSSFWorkbook wb1 = new XSSFWorkbook(new FileInputStream(BBH_2G));
			XSSFSheet BBH_2G_Sheet = wb1.getSheetAt(0);
			
			NBH_2G=new File(path+"NQI_NBH_REPORT_UPW_new.xlsx");
			XSSFWorkbook wb2 = new XSSFWorkbook(new FileInputStream(NBH_2G));
			XSSFSheet NBH_2G_Sheet = wb2.getSheetAt(0);
			
			NBH_3G=new File(path+"3G_KPI's.xlsx");
			XSSFWorkbook wb3 = new XSSFWorkbook(new FileInputStream(NBH_3G));
			XSSFSheet NBH_3G_Sheet = wb3.getSheetAt(0);
			
			//Write Data Start Here
			 XSSFWorkbook dworkbook=new XSSFWorkbook();
			 Sheet s1=dworkbook.createSheet("2G");
			 Cell ocell=null;
			 Row orow=null;
			
			 CellStyle newstyle5 = dworkbook.createCellStyle();
			 newstyle5.setFillForegroundColor(IndexedColors.BLUE.getIndex());
			 newstyle5.setFillPattern(CellStyle.SOLID_FOREGROUND);
			 newstyle5.setBorderBottom(CellStyle.BORDER_THIN);
			 newstyle5.setBorderLeft(CellStyle.BORDER_THIN);
			 newstyle5.setBorderRight(CellStyle.BORDER_THIN);
			 newstyle5.setBorderTop(CellStyle.BORDER_THIN);
			 newstyle5.setAlignment(CellStyle.ALIGN_CENTER);
			 XSSFFont  font = dworkbook.createFont();
			 font.setBold(true);
			 font.setColor(IndexedColors.WHITE.getIndex());
			 newstyle5.setFont(font);
			 
			 CellStyle borderStyle = dworkbook.createCellStyle();
			 borderStyle.setBorderBottom(CellStyle.BORDER_THIN);
			 borderStyle.setBorderLeft(CellStyle.BORDER_THIN);
			 borderStyle.setBorderRight(CellStyle.BORDER_THIN);
			 borderStyle.setBorderTop(CellStyle.BORDER_THIN);
			 borderStyle.setAlignment(CellStyle.ALIGN_CENTER);
			
			 DateFormat dateFormat = new SimpleDateFormat("dd-MMM");
			 Date date = new Date();
			 String date1=dateFormat.format(date);
			 System.out.println(date1);
			 
			 orow=s1.createRow(0);
			 ocell=orow.createCell(0);
			 ocell.setCellValue("Cell");
			 ocell.setCellStyle(newstyle5);
			 ocell=orow.createCell(1);
			 ocell.setCellValue("KPI Effected");
			 ocell.setCellStyle(newstyle5);
			 ocell=orow.createCell(2);
			 ocell.setCellValue("Target Value");
			 ocell.setCellStyle(newstyle5);
			 ocell=orow.createCell(3);
			 ocell.setCellValue(date1);
			 ocell.setCellStyle(newstyle5);
			 
			 Sheet s2=dworkbook.createSheet("3G");
			 Cell ocell1=null;
			 Row orow1=null;
			 
			 orow1=s2.createRow(0);
			 ocell1=orow1.createCell(0);
			 ocell1.setCellValue("Cell");
			 ocell1.setCellStyle(newstyle5);
			 ocell1=orow1.createCell(1);
			 ocell1.setCellValue("KPI Effected");
			 ocell1.setCellStyle(newstyle5);
			 ocell1=orow1.createCell(2);
			 ocell1.setCellValue("Target Value");
			 ocell1.setCellStyle(newstyle5);
			 ocell1=orow1.createCell(3);
			 ocell1.setCellValue(date1);
			 ocell1.setCellStyle(newstyle5);
			 
			 
			 Sheet s3=dworkbook.createSheet("Cell Count");
			 Cell ocell2=null;
			 Row orow2=null;
			 orow2=s3.createRow(0);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("KPIs");
			 ocell2.setCellStyle(newstyle5);
			 
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(date1);
			 ocell2.setCellStyle(newstyle5);
			 
			 orow2=s3.createRow(1);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value(DBBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(2);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value(DBBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(3);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value (DBBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(4);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value(DBBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(5);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("Handover Success Rate (BBH) New kpi");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(6);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("Handover Success Rate (NBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(7);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("RX Quality DL (0-5) (BBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 
			 orow2=s3.createRow(8);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("RX Quality DL (0-5) (NBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 
			 orow2=s3.createRow(9);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("SDCCH Assignment Success (BBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 
			 orow2=s3.createRow(10);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("SDCCH Assignment Success (NBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(11);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("SDCCH Completion Rate (BBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(12);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("SDCCH Completion Rate (NBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(13);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("TCH Assignment Success (BBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(14);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("TCH Assignment Success (NBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(15);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("TCH Completion Rate (BBH)");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(16);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("TCH Completion Rate NBH");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(17);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("DCR_PS");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(18);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("DCR_CS");
			 ocell2.setCellStyle(borderStyle);
			 
			 
			 orow2=s3.createRow(19);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("DSSR");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(20);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("RAB_Succ_PS_Rate");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(21);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("RAB_Succ_CS_Rate");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(22);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("CSSR");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(23);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("RRC_Succ_CS_Rate");
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.createRow(24);
			 ocell2=orow2.createCell(0);
			 ocell2.setCellValue("RRC_Succ_PS_Rate");
			 ocell2.setCellStyle(borderStyle);
			 
			 
			int Count_ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value1=0;
			int Count_ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value1=0;
			int Count_ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value1=0;
			int Count_ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value1=0;
			for(int i=0;i<DBBH_2G_Sheet.getPhysicalNumberOfRows();i++)
			{
				Row row=DBBH_2G_Sheet.getRow(i);
				if(row!=null)
				{
					if(i==0)
					{
						for(int P=0; P<row.getPhysicalNumberOfCells();P++)
						{
							Cell c2=row.getCell(P);
							if(c2!=null)
							{
								String ColHeading1=c2.getStringCellValue().trim();
								if(ColHeading1.equals("Short name"))
								{
									Short_name_Index=P;
								}
								else if(ColHeading1.equals("ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF BH Value"))
								{
									ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_Index=P;
								}
								else if(ColHeading1.equals("ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF BH Value"))
								{
									ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_Index=P;
								}
								else if(ColHeading1.equals("ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1 BH Value"))
								{
									ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_Index=P;
								}
								else if(ColHeading1.equals("ERCS_BSS_Dash_DL_Hard_Blocking BH Value"))
								{
									ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_Index=P;
								}		
							}
						}	
					}
					Cell Short_name=null;
					Cell ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value=null;
					Cell ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value=null;
					Cell ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value=null;
					Cell ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value=null;
					Short_name=row.getCell(Short_name_Index);
					ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value=row.getCell(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_Index);
					ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value=row.getCell(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_Index);
					ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value=row.getCell(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_Index);
					ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value=row.getCell(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_Index);
					if(i>=1)
					{
						String Short_name_Value="";
						if(Short_name!=null)
						{
							if(Short_name.getCellType()==Cell.CELL_TYPE_STRING)
							{
								Short_name_Value=Short_name.getStringCellValue();
							}
							else if(Short_name.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Short_name_Value=String.valueOf((int)Short_name.getNumericCellValue());
							}	
						}
						double ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value=0;
						if(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value!=null)
						{  
							if(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value=ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value.getNumericCellValue();
							
							}
							else if(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value =Double.parseDouble(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value = 0.0;
								}
							}
							else
							{
								ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value = 0.0;
							}
						}
						double ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value=0;
						if(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value!=null)
						{  
							if(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value=ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value.getNumericCellValue();
							
							}
							else if(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value =Double.parseDouble(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value = 0.0;
								}
							}
							else
							{
								ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value = 0.0;
							}
						}
						double ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value=0;
						if(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value!=null)
						{  
							if(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value=ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value.getNumericCellValue();
							
							}
							else if(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value =Double.parseDouble(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value = 0.0;
								}
							}
							else
							{
								ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value = 0.0;
							}
						}
						double ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value=0;
						if(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value!=null)
						{  
							if(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value=ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value.getNumericCellValue();
							
							}
							else if(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value =Double.parseDouble(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value = 0.0;
								}
							}
							else
							{
								ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value = 0.0;
							}
						}
						
						if(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value<25)
						{
							if(Short_name_Value.trim().length()>0)
							{
								Count_ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value1++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Short_name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<25");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value(DBBH)");
								ocell.setCellStyle(borderStyle);
								
								
							}
						}
						
						if(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value<11.25)
						{
							if(Short_name_Value.trim().length()>0)
							{
								Count_ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value1++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Short_name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<11.25");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value (DBBH)");
								ocell.setCellStyle(borderStyle);
								
							}	
						}
					
						if(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value<95)
						{
							if(Short_name_Value.trim().length()>0)
							{
								Count_ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value1++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Short_name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<95");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value(DBBH)");
								ocell.setCellStyle(borderStyle);
								
								
							}	
						}
						
						if(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value>2)
						{
							if(Short_name_Value.trim().length()>0)
							{
								Count_ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value1++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Short_name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue(">2");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value(DBBH)");
								ocell.setCellStyle(borderStyle);
								
								
							}
						}
					}
				}
			}
			
			//
			
			int Count_Handover_Success_Rate_BBH=0;
			int Count_SDCCH_Assignment_Success_BBH=0;
			int Count_SDCCH_Completion_Rate_BBH=0;
			int Count_TCH_Assignment_Success_BBH=0;
			int Count_TCH_Completion_Rate_BBH=0;
			int Count_RX_Quality_DL_BBH=0;
			
			for(int i=0; i<BBH_2G_Sheet.getPhysicalNumberOfRows();i++)
			{
				Row row=BBH_2G_Sheet.getRow(i);
				if(row!=null)
				{
					if(i==7)
					{
						for(int k=0; k<row.getPhysicalNumberOfCells();k++)
						{
							Cell c2=row.getCell(k);
							if(c2!=null)
							{
								String ColHeading1=c2.getStringCellValue().trim();
								if(ColHeading1.equals("Cell Name"))
								{
									Cell_Name_Index=k;
								}
								else if(ColHeading1.equals("VFE _HSR %"))
								{
									Handover_Success_Rate_Index=k;
								}
								else if(ColHeading1.equals("VFE_SDCCH BLK%"))
								{
									SDCCH_Assignment_Success_Index=k;
								}
								else if(ColHeading1.equals("TCH BLK BBH% VFE"))
								{
									TCH_Assignment_Success_Index=k;
								}
								else if(ColHeading1.equals("SD_drop (%)"))
								{
									SDCCH_Completion_Rate_Index=k;
								}
								else if(ColHeading1.equals("TCH_Drop %"))
								{
									TCH_Completion_Rate_Index=k;
								}
								else if(ColHeading1.equals("VFE TASR%"))
								{
									RX_Quality_DL_Index=k;
								}
							}
						}	
					}
					Cell Cell_Name=null;
					Cell Handover_Success_Rate=null;
					Cell SDCCH_Assignment_Success=null;
					Cell TCH_Assignment_Success=null;
					Cell SDCCH_Completion_Rate=null;
					Cell TCH_Completion_Rate=null;
					Cell RX_Quality_DL=null;
					Cell_Name=row.getCell(Cell_Name_Index);
					Handover_Success_Rate=row.getCell(Handover_Success_Rate_Index);
					SDCCH_Assignment_Success=row.getCell(SDCCH_Assignment_Success_Index);
					TCH_Assignment_Success=row.getCell(TCH_Assignment_Success_Index);
					SDCCH_Completion_Rate=row.getCell(SDCCH_Completion_Rate_Index);
					TCH_Completion_Rate=row.getCell(TCH_Completion_Rate_Index);
					RX_Quality_DL=row.getCell(RX_Quality_DL_Index);
					if(i>=8)
					{
						String Cell_Name_Value="";
						if(Cell_Name!=null)
						{
							if(Cell_Name.getCellType()==Cell.CELL_TYPE_STRING)
							{
								Cell_Name_Value=Cell_Name.getStringCellValue();
							}
							else if(Cell_Name.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Cell_Name_Value=String.valueOf((int)Cell_Name.getNumericCellValue());
							}	
						}
						double Handover_Success_Rate_value=0;
						if(Handover_Success_Rate!=null)
						{  
							if(Handover_Success_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Handover_Success_Rate_value=Handover_Success_Rate.getNumericCellValue();
							
							}
							else if(Handover_Success_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									Handover_Success_Rate_value =Double.parseDouble(Handover_Success_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									Handover_Success_Rate_value = 0.0;
								}
							}
							else
							{
								Handover_Success_Rate_value = 0.0;
							}
						}
						
						if(Handover_Success_Rate_value<92)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_Handover_Success_Rate_BBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<92");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(Handover_Success_Rate_value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("Handover Success Rate (BBH) New kpi");
								ocell.setCellStyle(borderStyle);
								
								
							}	
						}
						double SDCCH_Assignment_Success_value=0;
						if(SDCCH_Assignment_Success!=null)
						{  
							if(SDCCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								SDCCH_Assignment_Success_value=SDCCH_Assignment_Success.getNumericCellValue();
							
							}
							else if(SDCCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									SDCCH_Assignment_Success_value =Double.parseDouble(SDCCH_Assignment_Success.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									SDCCH_Assignment_Success_value = 0.0;
								}
							}
							else
							{
								SDCCH_Assignment_Success_value = 0.0;
							}
						}
						
						double SDCCH_Assignment_Success_value1=100.0-SDCCH_Assignment_Success_value;
						if(SDCCH_Assignment_Success_value1<99.5)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_SDCCH_Assignment_Success_BBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<99.5");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(SDCCH_Assignment_Success_value1);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("SDCCH Assignment Success (BBH)");
								ocell.setCellStyle(borderStyle);
								
							}
						}
						double TCH_Assignment_Success_value=0;
						if(TCH_Assignment_Success!=null)
						{  
							if(TCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								TCH_Assignment_Success_value=TCH_Assignment_Success.getNumericCellValue();
							
							}
							else if(TCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									TCH_Assignment_Success_value =Double.parseDouble(TCH_Assignment_Success.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									TCH_Assignment_Success_value = 0.0;
								}
							}
							else
							{
								TCH_Assignment_Success_value = 0.0;
							}
						}
						
						double TCH_Assignment_Success_value1=100.0-TCH_Assignment_Success_value;
						if(TCH_Assignment_Success_value1<98)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_TCH_Assignment_Success_BBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<98");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(TCH_Assignment_Success_value1);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("TCH Assignment Success (BBH)");
								ocell.setCellStyle(borderStyle);
								
							}
						}
						
						double SDCCH_Completion_Rate_value=0;
						if(SDCCH_Completion_Rate!=null)
						{  
							if(SDCCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								SDCCH_Completion_Rate_value=SDCCH_Completion_Rate.getNumericCellValue();
							
							}
							else if(SDCCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									SDCCH_Completion_Rate_value =Double.parseDouble(SDCCH_Completion_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									SDCCH_Completion_Rate_value = 0.0;
								}
							}
							else
							{
								SDCCH_Completion_Rate_value = 0.0;
							}
						}
						
						double SDCCH_Completion_Rate_value1=100.0-SDCCH_Completion_Rate_value;
						if(SDCCH_Completion_Rate_value1<98)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_SDCCH_Completion_Rate_BBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<98");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(SDCCH_Completion_Rate_value1);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("SDCCH Completion Rate (BBH)");
								ocell.setCellStyle(borderStyle);
								
							}
						}
						double TCH_Completion_Rate_value=0;
						if(TCH_Completion_Rate!=null)
						{  
							if(TCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								TCH_Completion_Rate_value=TCH_Completion_Rate.getNumericCellValue();
							
							}
							else if(TCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									TCH_Completion_Rate_value =Double.parseDouble(TCH_Completion_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									TCH_Completion_Rate_value = 0.0;
								}
							}
							else
							{
								TCH_Completion_Rate_value = 0.0;
							}
						}
					
						double TCH_Completion_Rate_value1=100.0-TCH_Completion_Rate_value;
						if(TCH_Completion_Rate_value1<98)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_TCH_Completion_Rate_BBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<98");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(TCH_Completion_Rate_value1);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("TCH Completion Rate (BBH)");
								ocell.setCellStyle(borderStyle);
								
							}	
						}
						double RX_Quality_DL_value=0;
						if(RX_Quality_DL!=null)
						{  
							if(RX_Quality_DL.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RX_Quality_DL_value=RX_Quality_DL.getNumericCellValue();
							
							}
							else if(RX_Quality_DL.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RX_Quality_DL_value =Double.parseDouble(RX_Quality_DL.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RX_Quality_DL_value = 0.0;
								}
							}
							else
							{
								RX_Quality_DL_value = 0.0;
							}
						}
					
						if(RX_Quality_DL_value<96)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_RX_Quality_DL_BBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<96");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(RX_Quality_DL_value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("RX Quality DL (0-5) (BBH)");
								ocell.setCellStyle(borderStyle);
								
							}
						}
					}		
				}
			}
			
			int Count_Handover_Success_Rate_NBH=0;
			int Count_RX_Quality_DL_NBH=0;
			int Count_SDCCH_Assignment_Success_NBH=0;
			int Count_SDCCH_Completion_Rate_NBH=0;
			int Count_TCH_Assignment_Success_NBH=0;
			int Count_TCH_Completion_Rate_NBH=0;
			
			for(int i=0;i<NBH_2G_Sheet.getPhysicalNumberOfRows();i++)
			{
				Row row=NBH_2G_Sheet.getRow(i);
				if(row!=null)
				{
					if(i==7)
					{
						for(int k=0; k<row.getPhysicalNumberOfCells();k++)
						{
							Cell c2=row.getCell(k);
							if(c2!=null)
							{
								String ColHeading1=c2.getStringCellValue().trim();
								if(ColHeading1.equals("Cell Name"))
								{
									Cell_Name_Index1=k;
								}
								else if(ColHeading1.equals("SD_drop (%)"))
								{
									SDCCH_Completion_Rate_Index1=k;
								}
								else if(ColHeading1.equals("TCH_Drop %"))
								{
									TCH_Completion_Rate_Index1=k;
								}
								else if(ColHeading1.equals("VFE _HSR %"))
								{
									Handover_Success_Rate_Index1=k;
								}
								else if(ColHeading1.equals("VFE_SDCCH BLK%"))
								{
									SDCCH_Assignment_Success_Index1=k;
								}
								else if(ColHeading1.equals("TCH BLK BBH% VFE"))
								{
									TCH_Assignment_Success_Index1=k;
								}
								else if(ColHeading1.equals("DL_QL"))
								{
									DL_QL_Index1=k;
								}
							}
						}
					}
					Cell Cell_Name=null;
					Cell SDCCH_Completion_Rate=null;
					Cell TCH_Completion_Rate=null;
					Cell Handover_Success_Rate=null;
					Cell SDCCH_Assignment_Success=null;
					Cell TCH_Assignment_Success=null;
					Cell DL_QL=null;
					Cell_Name=row.getCell(Cell_Name_Index1);
					SDCCH_Completion_Rate=row.getCell(SDCCH_Completion_Rate_Index1);
					TCH_Completion_Rate=row.getCell(TCH_Completion_Rate_Index1);
					Handover_Success_Rate=row.getCell(Handover_Success_Rate_Index1);
					SDCCH_Assignment_Success=row.getCell(SDCCH_Assignment_Success_Index1);
					TCH_Assignment_Success=row.getCell(TCH_Assignment_Success_Index1);
					DL_QL=row.getCell(DL_QL_Index1);
					if(i>=8)
					{
						String Cell_Name_Value="";
						if(Cell_Name!=null)
						{
							if(Cell_Name.getCellType()==Cell.CELL_TYPE_STRING)
							{
								Cell_Name_Value=Cell_Name.getStringCellValue();
							}
							else if(Cell_Name.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Cell_Name_Value=String.valueOf((int)Cell_Name.getNumericCellValue());
							}	
						}
						
						double SDCCH_Completion_Rate_value=0;
						if(SDCCH_Completion_Rate!=null)
						{  
							if(SDCCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								SDCCH_Completion_Rate_value=SDCCH_Completion_Rate.getNumericCellValue();
							
							}
							else if(SDCCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									SDCCH_Completion_Rate_value =Double.parseDouble(SDCCH_Completion_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									SDCCH_Completion_Rate_value = 0.0;
								}
							}
							else
							{
								SDCCH_Completion_Rate_value = 0.0;
							}
						}
						
						double SDCCH_Completion_Rate_value_value1=100.0-SDCCH_Completion_Rate_value;
						if(SDCCH_Completion_Rate_value_value1<98.8)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_SDCCH_Completion_Rate_NBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<98.8");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(SDCCH_Completion_Rate_value_value1);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("SDCCH Completion Rate (NBH)");
								ocell.setCellStyle(borderStyle);
								
							}
						}
						double TCH_Completion_Rate_value=0;
						if(TCH_Completion_Rate!=null)
						{  
							if(TCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								TCH_Completion_Rate_value=TCH_Completion_Rate.getNumericCellValue();
							
							}
							else if(TCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									TCH_Completion_Rate_value =Double.parseDouble(TCH_Completion_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									TCH_Completion_Rate_value = 0.0;
								}
							}
							else
							{
								TCH_Completion_Rate_value = 0.0;
							}
						}
						
						double TCH_Completion_Rate_value1=100.0-TCH_Completion_Rate_value;
						if(TCH_Completion_Rate_value1<98.5)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_TCH_Completion_Rate_NBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<98.5");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(TCH_Completion_Rate_value1);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("TCH Completion Rate NBH)");
								ocell.setCellStyle(borderStyle);
								
							}
						}
						double Handover_Success_Rate_value=0;
						if(Handover_Success_Rate!=null)
						{  
							if(Handover_Success_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Handover_Success_Rate_value=Handover_Success_Rate.getNumericCellValue();
							
							}
							else if(Handover_Success_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									Handover_Success_Rate_value =Double.parseDouble(Handover_Success_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									Handover_Success_Rate_value = 0.0;
								}
							}
							else
							{
								Handover_Success_Rate_value = 0.0;
							}
						}
						
						if(Handover_Success_Rate_value<96)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_Handover_Success_Rate_NBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<96");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(Handover_Success_Rate_value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("Handover Success Rate (NBH)");
								ocell.setCellStyle(borderStyle);
								
							}	
						}
						double SDCCH_Assignment_Success_value=0;
						if(SDCCH_Assignment_Success!=null)
						{  
							if(SDCCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								SDCCH_Assignment_Success_value=SDCCH_Assignment_Success.getNumericCellValue();
							
							}
							else if(SDCCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									SDCCH_Assignment_Success_value =Double.parseDouble(SDCCH_Assignment_Success.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									SDCCH_Assignment_Success_value = 0.0;
								}
							}
							else
							{
								SDCCH_Assignment_Success_value = 0.0;
							}
						}
						
						double SDCCH_Assignment_Success_value1=100.0-SDCCH_Assignment_Success_value;
						if(SDCCH_Assignment_Success_value1<99.6)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_SDCCH_Assignment_Success_NBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<99.6");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(SDCCH_Assignment_Success_value1);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("SDCCH Assignment Success (NBH)");
								ocell.setCellStyle(borderStyle);
								
							}
						}
						double TCH_Assignment_Success_value=0;
						if(TCH_Assignment_Success!=null)
						{  
							if(TCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								TCH_Assignment_Success_value=TCH_Assignment_Success.getNumericCellValue();
							
							}
							else if(TCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									TCH_Assignment_Success_value =Double.parseDouble(TCH_Assignment_Success.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									TCH_Assignment_Success_value = 0.0;
								}
							}
							else
							{
								TCH_Assignment_Success_value = 0.0;
							}
						}
					
						double TCH_Assignment_Success_value1=100.0-TCH_Assignment_Success_value;
						if(TCH_Assignment_Success_value1<99)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_TCH_Assignment_Success_NBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<99");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(TCH_Assignment_Success_value1);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("TCH Assignment Success (NBH)");
								ocell.setCellStyle(borderStyle);
							
							}
						}
						double RX_Quality_DL_value=0;
						if(DL_QL!=null)
						{  
							if(DL_QL.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RX_Quality_DL_value=DL_QL.getNumericCellValue();
							
							}
							else if(DL_QL.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RX_Quality_DL_value =Double.parseDouble(DL_QL.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RX_Quality_DL_value = 0.0;
								}
							}
							else
							{
								RX_Quality_DL_value = 0.0;
							}
						}
						
						if(RX_Quality_DL_value<97)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_RX_Quality_DL_NBH++;
								orow=s1.createRow(s1.getPhysicalNumberOfRows());
								ocell=orow.createCell(0);
								ocell.setCellValue(Cell_Name_Value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(2);
								ocell.setCellValue("<97");
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(3);
								ocell.setCellValue(RX_Quality_DL_value);
								ocell.setCellStyle(borderStyle);
								ocell=orow.createCell(1);
								ocell.setCellValue("RX Quality DL (0-5) (NBH)");
								ocell.setCellStyle(borderStyle);
							
							}
						}
					}
				}							
			}
			
			int Count_DCR_PS=0;
			int Count_DCR_CS=0;
			int Count_DSSR=0;
			int Count_RAB_Succ_PS_Rate=0;
			int Count_RAB_Succ_CS_Rate=0;
			int Count_CSSR=0;
			int Count_RRC_Succ_CS_Rate=0;
			int Count_RRC_Succ_PS_Rate=0;
			for(int i=0;i<NBH_3G_Sheet.getPhysicalNumberOfRows();i++)
			{
				Row row=NBH_3G_Sheet.getRow(i);
				if(row!=null)
				{
					if(i==3)
					{
						for(int k=0; k<row.getPhysicalNumberOfCells();k++)
						{
							Cell c2=row.getCell(k);
							if(c2!=null)
							{
								String ColHeading1=c2.getStringCellValue().trim();
								if(ColHeading1.equals("UtranCell"))
								{
									UtranCell_Index=k;
								}
								else if(ColHeading1.equals("DCR_PS"))
								{
									DCR_PS_Index=k;
								}
								else if(ColHeading1.equals("DCR_CS"))
								{
									DCR_CS_Index=k;
								}
								else if(ColHeading1.equals("RRC_Succ_CS_Rate"))
								{
									RRC_Succ_CS_Rate_Index=k;
								}
								else if(ColHeading1.equals("RRC_Succ_PS_Rate"))
								{
									RRC_Succ_PS_Rate_Index=k;
								}
								else if(ColHeading1.equals("RAB_Succ_CS_Rate"))
								{
									RAB_Succ_CS_Rate_Index=k;
								}
								else if(ColHeading1.equals("RAB_Succ_PS_Rate"))
								{
									RAB_Succ_PS_Rate_Index=k;
								}
								else if(ColHeading1.equals("CSSR"))
								{
									CSSR_Index=k;
								}
								else if(ColHeading1.equals("DSSR"))
								{
									DSSR_Index=k;
								}
							}
						}
					}
					Cell UtranCell=null;
					Cell DCR_PS=null;
					Cell DCR_CS=null;
					Cell RRC_Succ_CS_Rate=null;
					Cell RRC_Succ_PS_Rate=null;
					Cell RAB_Succ_CS_Rate=null;
					Cell RAB_Succ_PS_Rate=null;
					Cell CSSR=null;
					Cell DSSR=null;
					UtranCell=row.getCell(UtranCell_Index);
					DCR_PS=row.getCell(DCR_PS_Index);
					DCR_CS=row.getCell(DCR_CS_Index);
					RRC_Succ_CS_Rate=row.getCell(RRC_Succ_CS_Rate_Index);
					RRC_Succ_PS_Rate=row.getCell(RRC_Succ_PS_Rate_Index);	
					RAB_Succ_CS_Rate=row.getCell(RAB_Succ_CS_Rate_Index);
					RAB_Succ_PS_Rate=row.getCell(RAB_Succ_PS_Rate_Index);
					CSSR=row.getCell(CSSR_Index);
					DSSR=row.getCell(DSSR_Index);
					if(i>=4)
					{
						String UtranCell_Value="";
						if(UtranCell!=null)
						{
							if(UtranCell.getCellType()==Cell.CELL_TYPE_STRING)
							{
								UtranCell_Value=UtranCell.getStringCellValue();
							}
							else if(UtranCell.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								UtranCell_Value=String.valueOf((int)UtranCell.getNumericCellValue());
							}	
						}
						double DCR_PS_value=0;
						if(DCR_PS!=null)
						{  
							if(DCR_PS.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								DCR_PS_value=DCR_PS.getNumericCellValue();
							}
							else if(DCR_PS.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									DCR_PS_value =Double.parseDouble(DCR_PS.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									DCR_PS_value = 0.0;
								}
							}
							else
							{
								DCR_PS_value = 0.0;
							}
						}
					
						if(DCR_PS_value>2)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_DCR_PS++;
								orow1=s2.createRow(s2.getPhysicalNumberOfRows());
								ocell1=orow1.createCell(0);
								ocell1.setCellValue(UtranCell_Value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(2);
								ocell1.setCellValue(">2");
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(3);
								ocell1.setCellValue(DCR_PS_value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(1);
								ocell1.setCellValue("DCR_PS");
								ocell1.setCellStyle(borderStyle);
								
							}
						}
						double DCR_CS_value=0;
						if(DCR_CS!=null)
						{  
							if(DCR_CS.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								DCR_CS_value=DCR_CS.getNumericCellValue();
							
							}
							else if(DCR_CS.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									DCR_CS_value =Double.parseDouble(DCR_CS.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									DCR_CS_value = 0.0;
								}
							}
							else
							{
								DCR_CS_value = 0.0;
							}
						}
						
						if(DCR_CS_value>2)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_DCR_CS++;
								orow1=s2.createRow(s2.getPhysicalNumberOfRows());
								ocell1=orow1.createCell(0);
								ocell1.setCellValue(UtranCell_Value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(2);
								ocell1.setCellValue(">2");
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(3);
								ocell1.setCellValue(DCR_CS_value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(1);
								ocell1.setCellValue("DCR_CS");
								ocell1.setCellStyle(borderStyle);
							}	
						}
						double RRC_Succ_CS_Rate_value=0;
						if(RRC_Succ_CS_Rate!=null)
						{  
							if(RRC_Succ_CS_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RRC_Succ_CS_Rate_value=RRC_Succ_CS_Rate.getNumericCellValue();
							}
							else if(RRC_Succ_CS_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RRC_Succ_CS_Rate_value =Double.parseDouble(RRC_Succ_CS_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RRC_Succ_CS_Rate_value = 0.0;
								}
							}
							else
							{
								RRC_Succ_CS_Rate_value = 0.0;
							}
						}
						
						if(RRC_Succ_CS_Rate_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_RRC_Succ_CS_Rate++;
								orow1=s2.createRow(s2.getPhysicalNumberOfRows());
								ocell1=orow1.createCell(0);
								ocell1.setCellValue(UtranCell_Value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(2);
								ocell1.setCellValue("<99");
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(3);
								ocell1.setCellValue(RRC_Succ_CS_Rate_value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(1);
								ocell1.setCellValue("RRC_Succ_CS_Rate");
								ocell1.setCellStyle(borderStyle);
							
							}
						}
						double RRC_Succ_PS_Rate_value=0;
						if(RRC_Succ_PS_Rate!=null)
						{  
							if(RRC_Succ_PS_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RRC_Succ_PS_Rate_value=RRC_Succ_PS_Rate.getNumericCellValue();
							
							}
							else if(RRC_Succ_PS_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RRC_Succ_PS_Rate_value =Double.parseDouble(RRC_Succ_PS_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RRC_Succ_PS_Rate_value = 0.0;
								}
							}
							else
							{
								RRC_Succ_PS_Rate_value = 0.0;
							}
						}
						
						if(RRC_Succ_PS_Rate_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_RRC_Succ_PS_Rate++;
								orow1=s2.createRow(s2.getPhysicalNumberOfRows());
								ocell1=orow1.createCell(0);
								ocell1.setCellValue(UtranCell_Value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(2);
								ocell1.setCellValue("<99");
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(3);
								ocell1.setCellValue(RRC_Succ_PS_Rate_value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(1);
								ocell1.setCellValue("RRC_Succ_PS_Rate");
								ocell1.setCellStyle(borderStyle);
								
							}
						}
						
						double RAB_Succ_CS_Rate_value=0;
						if(RAB_Succ_CS_Rate!=null)
						{  
							if(RAB_Succ_CS_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RAB_Succ_CS_Rate_value=RAB_Succ_CS_Rate.getNumericCellValue();
							
							}
							else if(RAB_Succ_CS_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RAB_Succ_CS_Rate_value =Double.parseDouble(RAB_Succ_CS_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RAB_Succ_CS_Rate_value = 0.0;
								}
							}
							else
							{
								RAB_Succ_CS_Rate_value = 0.0;
							}
						}	
						
						if(RAB_Succ_CS_Rate_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_RAB_Succ_CS_Rate++;
								orow1=s2.createRow(s2.getPhysicalNumberOfRows());
								ocell1=orow1.createCell(0);
								ocell1.setCellValue(UtranCell_Value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(2);
								ocell1.setCellValue("<99");
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(3);
								ocell1.setCellValue(RAB_Succ_CS_Rate_value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(1);
								ocell1.setCellValue("RAB_Succ_CS_Rate");
								ocell1.setCellStyle(borderStyle);
								
							}
						}
						double RAB_Succ_PS_Rate_value=0;
						if(RAB_Succ_PS_Rate!=null)
						{  
							if(RAB_Succ_PS_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RAB_Succ_PS_Rate_value=RAB_Succ_PS_Rate.getNumericCellValue();
							}
							else if(RAB_Succ_PS_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RAB_Succ_PS_Rate_value =Double.parseDouble(RAB_Succ_PS_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RAB_Succ_PS_Rate_value = 0.0;
								}
							}
							else
							{
								RAB_Succ_PS_Rate_value = 0.0;
							}
						}	
					
						if(RAB_Succ_PS_Rate_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_RAB_Succ_PS_Rate++;
								orow1=s2.createRow(s2.getPhysicalNumberOfRows());
								ocell1=orow1.createCell(0);
								ocell1.setCellValue(UtranCell_Value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(2);
								ocell1.setCellValue("<99");
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(3);
								ocell1.setCellValue(RAB_Succ_PS_Rate_value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(1);
								ocell1.setCellValue("RAB_Succ_PS_Rate");
								ocell1.setCellStyle(borderStyle);
								
							}
						}
						double CSSR_value=0;
						if(CSSR!=null)
						{  
							if(CSSR.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								CSSR_value=CSSR.getNumericCellValue();
							
							}
							else if(CSSR.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									CSSR_value =Double.parseDouble(CSSR.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									CSSR_value = 0.0;
								}
							}
							else
							{
								CSSR_value = 0.0;
							}
						}	
						
						if(CSSR_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_CSSR++;
								orow1=s2.createRow(s2.getPhysicalNumberOfRows());
								ocell1=orow1.createCell(0);
								ocell1.setCellValue(UtranCell_Value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(2);
								ocell1.setCellValue("<99");
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(3);
								ocell1.setCellValue(CSSR_value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(1);
								ocell1.setCellValue("CSSR");
								ocell1.setCellStyle(borderStyle);
								
							}
						}
						double DSSR_value=0;
						if(DSSR!=null)
						{  
							if(DSSR.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								DSSR_value=DSSR.getNumericCellValue();
							}
							else if(DSSR.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									DSSR_value =Double.parseDouble(DSSR.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									DSSR_value = 0.0;
								}
							}
							else
							{
								DSSR_value = 0.0;
							}
						}	
					
						if(DSSR_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_DSSR++;
								orow1=s2.createRow(s2.getPhysicalNumberOfRows());
								ocell1=orow1.createCell(0);
								ocell1.setCellValue(UtranCell_Value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(2);
								ocell1.setCellValue("<99");
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(3);
								ocell1.setCellValue(DSSR_value);
								ocell1.setCellStyle(borderStyle);
								ocell1=orow1.createCell(1);
								ocell1.setCellValue("DSSR");
								ocell1.setCellStyle(borderStyle);
								
							}
						}
					}
				}					
			}
			orow=s1.getRow(0);
			ocell=orow.createCell(orow.getPhysicalNumberOfCells());
			ocell.setCellValue("Count");
			ocell.setCellStyle(newstyle5);
			
			orow=s1.getRow(0);
			ocell=orow.createCell(orow.getPhysicalNumberOfCells());
			ocell.setCellValue("Ranking");
			ocell.setCellStyle(newstyle5);
			int cell_Count=orow.getPhysicalNumberOfCells()-2;
			int count=0;
			for(int i=1;i<s1.getPhysicalNumberOfRows();i++)
			{
				Row row=s1.getRow(i);
				if(row!=null)
				{
					count=0;
					for(int j=3;j<12;j++)
					{
						Cell rowwisedata=row.getCell(j);
						if(rowwisedata!=null)
						{
							count++;
						}
					}
					Cell cell=row.createCell(cell_Count);
					cell.setCellValue(count);
					cell.setCellStyle(borderStyle);
					if(count==10)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("1");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==9)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("2");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==8)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("3");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==7)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("4");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==6)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("5");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==5)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("6");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==4)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("7");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==3)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("8");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==2)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("9");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==1)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("10");
						cell1.setCellStyle(borderStyle);
					}
				}
			}
			orow1=s2.getRow(0);
			ocell1=orow1.createCell(orow1.getPhysicalNumberOfCells());
			ocell1.setCellValue("Count");
			ocell1.setCellStyle(newstyle5);
			
			orow1=s2.getRow(0);
			ocell1=orow1.createCell(orow1.getPhysicalNumberOfCells());
			ocell1.setCellValue("Ranking");
			ocell1.setCellStyle(newstyle5);
			int cell_Count1=orow.getPhysicalNumberOfCells()-2;
			int count1=0;
			
			for(int i=1;i<s2.getPhysicalNumberOfRows();i++)
			{
				Row row=s2.getRow(i);
				if(row!=null)
				{
					count1=0;
					for(int j=3;j<12;j++)
					{
						Cell rowwisedata=row.getCell(j);
						if(rowwisedata!=null)
						{
							count1++;
						}
					}
					Cell cell=row.createCell(cell_Count1);
					cell.setCellValue(count1);
					cell.setCellStyle(borderStyle);
					if(count1==10)
					{
						Cell cell1=row.createCell(cell_Count1+1);
						cell1.setCellValue("1");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==9)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("2");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==8)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("3");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==7)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("4");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==6)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("5");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==5)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("6");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==4)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("7");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==3)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("8");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==2)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("9");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==1)
					{
						Cell cell1=row.createCell(cell_Count+1);
						cell1.setCellValue("10");
						cell1.setCellStyle(borderStyle);
					}
				}
			}

			// Count of Cell Sheet Write Start here
			 orow2=s3.getRow(1);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value1);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(2);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value1);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(3);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value1);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(4);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value1);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(5);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_Handover_Success_Rate_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(6);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_Handover_Success_Rate_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(7);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_RX_Quality_DL_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(8);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_RX_Quality_DL_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(9);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_SDCCH_Assignment_Success_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(10);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_SDCCH_Assignment_Success_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(11);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_SDCCH_Completion_Rate_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(12);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_SDCCH_Completion_Rate_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(13);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_TCH_Assignment_Success_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(14);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_TCH_Assignment_Success_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(15);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_TCH_Completion_Rate_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(16);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_TCH_Completion_Rate_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(17);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_DCR_PS);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(18);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_DCR_CS);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(19);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_DSSR);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(20);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_RAB_Succ_PS_Rate);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(21);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_RAB_Succ_CS_Rate);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(22);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_CSSR);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(23);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_RRC_Succ_CS_Rate);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(24);
			 ocell2=orow2.createCell(1);
			 ocell2.setCellValue(Count_RRC_Succ_PS_Rate);
			 ocell2.setCellStyle(borderStyle);
			
			FileOutputStream outExcel=new FileOutputStream(new File(path+"Output Worst Cell.xlsx"));              
			dworkbook.write(outExcel); //write the output data
			dworkbook.close();
			outExcel.close();
			wb.close();
			wb1.close();
			wb2.close();
			wb3.close();	
		
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	public void NewCode() throws IOException
	{		
		
		Hash_Map_2G=new LinkedHashMap<String, Double>(); 
		Hash_Map_3G=new LinkedHashMap<String, Double>(); 
		List_2G=new LinkedList<String>();
		List_3G=new LinkedList<String>();
		System.out.println("Ok");
		String path="C:\\MY Work\\myPro\\SoftHumanActivity\\WebContent\\Excel\\Worst_Cell\\";
		File DBBH_2G1=null;
		File BBH_2G1=null;
		File NBH_2G1=null;
		File NBH_3G1=null;
		File f=new File(path+"Output Worst Cell.xlsx");
				
			DBBH_2G1=new File(path+"ERCS_BSS_CELL_WISE_CELL_DBBH_UPW_09122015.xlsx");
			XSSFWorkbook wbb = new XSSFWorkbook(new FileInputStream(DBBH_2G1));
			XSSFSheet DBBH_2G_Sheet1 = wbb.getSheetAt(0);
			
			BBH_2G1=new File(path+"NQI_BBH_REPORT_UPW_new.xlsx");
			XSSFWorkbook wbb1 = new XSSFWorkbook(new FileInputStream(BBH_2G1));
			XSSFSheet BBH_2G_Sheet1 = wbb1.getSheetAt(0);
			
			NBH_2G1=new File(path+"NQI_NBH_REPORT_UPW_new.xlsx");
			XSSFWorkbook wbb2 = new XSSFWorkbook(new FileInputStream(NBH_2G1));
			XSSFSheet NBH_2G_Sheet1 = wbb2.getSheetAt(0);
			
			NBH_3G1=new File(path+"3G_KPI's.xlsx");
			XSSFWorkbook wbb3 = new XSSFWorkbook(new FileInputStream(NBH_3G1));
			XSSFSheet NBH_3G_Sheet1 = wbb3.getSheetAt(0);
			
			XSSFWorkbook wbb4 = new XSSFWorkbook(new FileInputStream(f));
			XSSFSheet Final_2G = wbb4.getSheetAt(0);
			XSSFSheet Final_3G = wbb4.getSheetAt(1);
			
			for(int i=0; i<Final_2G.getPhysicalNumberOfRows();i++)
			{
				Row row= Final_2G.getRow(i);
				if(row!=null)
				{
					Cell cell_Name=row.getCell(0);
					Cell KPI_effected=row.getCell(1);
					String cell_Name_Value="";
					if(cell_Name!=null)
					{	
						if(cell_Name.getCellType()==Cell.CELL_TYPE_STRING)
						{
							cell_Name_Value=cell_Name.getStringCellValue();
						}	
					}
					String KPI_effected_Value="";
					if(KPI_effected!=null)
					{	
						if(KPI_effected.getCellType()==Cell.CELL_TYPE_STRING)
						{
							KPI_effected_Value=KPI_effected.getStringCellValue();
						}	
					}
					if(!cell_Name_Value.equals("") && !KPI_effected_Value.equals(""))
					{
						List_2G.add(cell_Name_Value+KPI_effected_Value);
					}
					
				}
			}
			
			
			
			
			for(int i=0; i<Final_3G.getPhysicalNumberOfRows();i++)
			{
				Row row= Final_3G.getRow(i);
				if(row!=null)
				{
					Cell cell_Name=row.getCell(0);
					Cell KPI_effected=row.getCell(1);
					String cell_Name_Value="";
					if(cell_Name!=null)
					{	
						if(cell_Name.getCellType()==Cell.CELL_TYPE_STRING)
						{
							cell_Name_Value=cell_Name.getStringCellValue();
						}	
					}
					String KPI_effected_Value="";
					if(KPI_effected!=null)
					{	
						if(KPI_effected.getCellType()==Cell.CELL_TYPE_STRING)
						{
							KPI_effected_Value=KPI_effected.getStringCellValue();
						}	
					}
					if(!cell_Name_Value.equals("") && !KPI_effected_Value.equals(""))
					{
						List_3G.add(cell_Name_Value+KPI_effected_Value);
					}
				}
			}
			 CellStyle newstyle5 = wbb4.createCellStyle();
			 newstyle5.setFillForegroundColor(IndexedColors.BLUE.getIndex());
			 newstyle5.setFillPattern(CellStyle.SOLID_FOREGROUND);
			 newstyle5.setBorderBottom(CellStyle.BORDER_THIN);
			 newstyle5.setBorderLeft(CellStyle.BORDER_THIN);
			 newstyle5.setBorderRight(CellStyle.BORDER_THIN);
			 newstyle5.setBorderTop(CellStyle.BORDER_THIN);
			 newstyle5.setAlignment(CellStyle.ALIGN_CENTER);
			 XSSFFont  font = wbb4.createFont();
			 font.setBold(true);
			 font.setColor(IndexedColors.WHITE.getIndex());
			 newstyle5.setFont(font);
			
			 
			 CellStyle borderStyle = wbb4.createCellStyle();
			 borderStyle.setBorderBottom(CellStyle.BORDER_THIN);
			 borderStyle.setBorderLeft(CellStyle.BORDER_THIN);
			 borderStyle.setBorderRight(CellStyle.BORDER_THIN);
			 borderStyle.setBorderTop(CellStyle.BORDER_THIN);
			 borderStyle.setAlignment(CellStyle.ALIGN_CENTER);
			
			 DateFormat dateFormat = new SimpleDateFormat("dd-MMM");
			 Date date2 = new Date();
			 String date3=dateFormat.format(date2);
			 System.out.println(date3);
			 
			 
			 Sheet s1=wbb4.getSheet("2G");
			 Cell ocell=null;
			 Row orow=null;
			 
			 Sheet s2=wbb4.getSheet("3G");
			 Cell ocell1=null;
			 Row orow1=null;
			 
			 Sheet s3=wbb4.getSheet("Cell Count");
			 Cell ocell2=null;
			 Row orow2=null;
			 
			 int Count_ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value1=0;
			 int Count_ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value1=0;
			 int Count_ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value1=0;
			 int Count_ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value1=0;
			
			for(int i=0;i<DBBH_2G_Sheet1.getPhysicalNumberOfRows();i++)
			{
				Row row=DBBH_2G_Sheet1.getRow(i);
				if(row!=null)
				{
					if(i==0)
					{
						for(int P=0; P<row.getPhysicalNumberOfCells();P++)
						{
							Cell c2=row.getCell(P);
							if(c2!=null)
							{
								String ColHeading1=c2.getStringCellValue().trim();
								if(ColHeading1.equals("Short name"))
								{
									Short_name_Index=P;
								}
								else if(ColHeading1.equals("ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF BH Value"))
								{
									ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_Index=P;
								}
								else if(ColHeading1.equals("ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF BH Value"))
								{
									ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_Index=P;
								}
								else if(ColHeading1.equals("ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1 BH Value"))
								{
									ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_Index=P;
								}
								else if(ColHeading1.equals("ERCS_BSS_Dash_DL_Hard_Blocking BH Value"))
								{
									ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_Index=P;
								}		
							}
						}	
					}
					Cell Short_name=null;
					Cell ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value=null;
					Cell ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value=null;
					Cell ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value=null;
					Cell ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value=null;
					Short_name=row.getCell(Short_name_Index);
					ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value=row.getCell(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_Index);
					ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value=row.getCell(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_Index);
					ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value=row.getCell(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_Index);
					ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value=row.getCell(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_Index);
					if(i>=1)
					{
						String Short_name_Value="";
						if(Short_name!=null)
						{
							if(Short_name.getCellType()==Cell.CELL_TYPE_STRING)
							{
								Short_name_Value=Short_name.getStringCellValue();
							}
							else if(Short_name.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Short_name_Value=String.valueOf((int)Short_name.getNumericCellValue());
							}	
						}
						double ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value=0;
						if(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value!=null)
						{  
							if(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value=ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value.getNumericCellValue();
							
							}
							else if(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value =Double.parseDouble(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value = 0.0;
								}
							}
							else
							{
								ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value = 0.0;
							}
						}
						double ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value=0;
						if(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value!=null)
						{  
							if(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value=ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value.getNumericCellValue();
							
							}
							else if(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value =Double.parseDouble(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value = 0.0;
								}
							}
							else
							{
								ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value = 0.0;
							}
						}
						double ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value=0;
						if(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value!=null)
						{  
							if(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value=ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value.getNumericCellValue();
							
							}
							else if(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value =Double.parseDouble(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value = 0.0;
								}
							}
							else
							{
								ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value = 0.0;
							}
						}
						double ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value=0;
						if(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value!=null)
						{  
							if(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value=ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value.getNumericCellValue();
							
							}
							else if(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value =Double.parseDouble(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value = 0.0;
								}
							}
							else
							{
								ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value = 0.0;
							}
						}
						if(ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value<25)
						{
							if(Short_name_Value.trim().length()>0)
							{
								Count_ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value1++;
								if(List_2G.contains(Short_name_Value+"ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value(DBBH)"))
								{
									Hash_Map_2G.put(Short_name_Value+"ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value(DBBH)", ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Short_name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<25");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value(DBBH)");
									ocell.setCellStyle(borderStyle);
									
									Hash_Map_2G.put(Short_name_Value+"ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value(DBBH)", ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value_value);
								}	
							}
						}
						if(ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value<11.25)
						{
							if(Short_name_Value.trim().length()>0)
							{
								Count_ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value1++;
								if(List_2G.contains(Short_name_Value+"ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value (DBBH)"))
								{
									Hash_Map_2G.put(Short_name_Value+"ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value (DBBH)", ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Short_name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<11.25");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value (DBBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Short_name_Value+"ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value (DBBH)", ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value_value);
								}
							}	
						}
						if(ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value<95)
						{
							if(Short_name_Value.trim().length()>0)
							{
								Count_ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value1++;
								if(List_2G.contains(Short_name_Value+"ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value(DBBH)"))
								{
									Hash_Map_2G.put(Short_name_Value+"ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value(DBBH)", ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Short_name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<95");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value(DBBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Short_name_Value+"ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value(DBBH)", ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value_value);
								}
							}	
						}
						if(ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value>2)
						{
							if(Short_name_Value.trim().length()>0)
							{
								Count_ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value1++;
								if(List_2G.contains(Short_name_Value+"ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value(DBBH)"))
								{
									Hash_Map_2G.put(Short_name_Value+"ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value(DBBH)", ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Short_name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue(">2");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value(DBBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Short_name_Value+"ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value(DBBH)", ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value_value);
								}
							}
						}
					}
				}
			}
			 int Count_Handover_Success_Rate_BBH=0;
			 int Count_SDCCH_Assignment_Success_BBH=0;
			 int Count_SDCCH_Completion_Rate_BBH=0;
			 int Count_TCH_Assignment_Success_BBH=0;
			 int Count_TCH_Completion_Rate_BBH=0;
			 int Count_RX_Quality_DL_BBH=0;
			
			for(int i=0; i<BBH_2G_Sheet1.getPhysicalNumberOfRows();i++)
			{
				Row row=BBH_2G_Sheet1.getRow(i);
				if(row!=null)
				{
					if(i==7)
					{
						for(int k=0; k<row.getPhysicalNumberOfCells();k++)
						{
							Cell c2=row.getCell(k);
							if(c2!=null)
							{
								String ColHeading1=c2.getStringCellValue().trim();
								if(ColHeading1.equals("Cell Name"))
								{
									Cell_Name_Index=k;
								}
								else if(ColHeading1.equals("VFE _HSR %"))
								{
									Handover_Success_Rate_Index=k;
								}
								else if(ColHeading1.equals("VFE_SDCCH BLK%"))
								{
									SDCCH_Assignment_Success_Index=k;
								}
								else if(ColHeading1.equals("TCH BLK BBH% VFE"))
								{
									TCH_Assignment_Success_Index=k;
								}
								else if(ColHeading1.equals("SD_drop (%)"))
								{
									SDCCH_Completion_Rate_Index=k;
								}
								else if(ColHeading1.equals("TCH_Drop %"))
								{
									TCH_Completion_Rate_Index=k;
								}
								else if(ColHeading1.equals("VFE TASR%"))
								{
									RX_Quality_DL_Index=k;
								}
							}
						}	
					}
					Cell Cell_Name=null;
					Cell Handover_Success_Rate=null;
					Cell SDCCH_Assignment_Success=null;
					Cell TCH_Assignment_Success=null;
					Cell SDCCH_Completion_Rate=null;
					Cell TCH_Completion_Rate=null;
					Cell RX_Quality_DL=null;
					Cell_Name=row.getCell(Cell_Name_Index);
					Handover_Success_Rate=row.getCell(Handover_Success_Rate_Index);
					SDCCH_Assignment_Success=row.getCell(SDCCH_Assignment_Success_Index);
					TCH_Assignment_Success=row.getCell(TCH_Assignment_Success_Index);
					SDCCH_Completion_Rate=row.getCell(SDCCH_Completion_Rate_Index);
					TCH_Completion_Rate=row.getCell(TCH_Completion_Rate_Index);
					RX_Quality_DL=row.getCell(RX_Quality_DL_Index);
					if(i>=8)
					{
						String Cell_Name_Value="";
						if(Cell_Name!=null)
						{
							if(Cell_Name.getCellType()==Cell.CELL_TYPE_STRING)
							{
								Cell_Name_Value=Cell_Name.getStringCellValue();
							}
							else if(Cell_Name.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Cell_Name_Value=String.valueOf((int)Cell_Name.getNumericCellValue());
							}	
						}
						double Handover_Success_Rate_value=0;
						if(Handover_Success_Rate!=null)
						{  
							if(Handover_Success_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Handover_Success_Rate_value=Handover_Success_Rate.getNumericCellValue();
							
							}
							else if(Handover_Success_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									Handover_Success_Rate_value =Double.parseDouble(Handover_Success_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									Handover_Success_Rate_value = 0.0;
								}
							}
							else
							{
								Handover_Success_Rate_value = 0.0;
							}
						}
						 
						
						if(Handover_Success_Rate_value<92)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_Handover_Success_Rate_BBH++;
								if(List_2G.contains(Cell_Name_Value+"Handover Success Rate (BBH) New kpi"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"Handover Success Rate (BBH) New kpi", Handover_Success_Rate_value);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<92");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("Handover Success Rate (BBH) New kpi");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"Handover Success Rate (BBH) New kpi", Handover_Success_Rate_value);
									
								}	
							}	
						}
						double SDCCH_Assignment_Success_value=0;
						if(SDCCH_Assignment_Success!=null)
						{  
							if(SDCCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								SDCCH_Assignment_Success_value=SDCCH_Assignment_Success.getNumericCellValue();
							
							}
							else if(SDCCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									SDCCH_Assignment_Success_value =Double.parseDouble(SDCCH_Assignment_Success.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									SDCCH_Assignment_Success_value = 0.0;
								}
							}
							else
							{
								SDCCH_Assignment_Success_value = 0.0;
							}
						}
						
						double SDCCH_Assignment_Success_value1=100.0-SDCCH_Assignment_Success_value;
						if(SDCCH_Assignment_Success_value1<99.5)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_SDCCH_Assignment_Success_BBH++;
								if(List_2G.contains(Cell_Name_Value+"SDCCH Assignment Success (BBH)"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"SDCCH Assignment Success (BBH)", SDCCH_Assignment_Success_value1);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<99.5");
									ocell.setCellStyle(borderStyle);
									
									ocell=orow.createCell(1);
									ocell.setCellValue("SDCCH Assignment Success (BBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"SDCCH Assignment Success (BBH)", SDCCH_Assignment_Success_value1);
								}	
							}
						}
						double TCH_Assignment_Success_value=0;
						if(TCH_Assignment_Success!=null)
						{  
							if(TCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								TCH_Assignment_Success_value=TCH_Assignment_Success.getNumericCellValue();
							
							}
							else if(TCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									TCH_Assignment_Success_value =Double.parseDouble(TCH_Assignment_Success.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									TCH_Assignment_Success_value = 0.0;
								}
							}
							else
							{
								TCH_Assignment_Success_value = 0.0;
							}
						}
						
						double TCH_Assignment_Success_value1=100.0-TCH_Assignment_Success_value;
						if(TCH_Assignment_Success_value1<98)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_TCH_Assignment_Success_BBH++;
								if(List_2G.contains(Cell_Name_Value+"TCH Assignment Success (BBH)"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"TCH Assignment Success (BBH)", TCH_Assignment_Success_value1);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<98");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("TCH Assignment Success (BBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"TCH Assignment Success (BBH)", TCH_Assignment_Success_value1);
									
								}
							}
						}
						double SDCCH_Completion_Rate_value=0;
						if(SDCCH_Completion_Rate!=null)
						{  
							if(SDCCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								SDCCH_Completion_Rate_value=SDCCH_Completion_Rate.getNumericCellValue();
							
							}
							else if(SDCCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									SDCCH_Completion_Rate_value =Double.parseDouble(SDCCH_Completion_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									SDCCH_Completion_Rate_value = 0.0;
								}
							}
							else
							{
								SDCCH_Completion_Rate_value = 0.0;
							}
						}
						 
						double SDCCH_Completion_Rate_value1=100.0-SDCCH_Completion_Rate_value;
						if(SDCCH_Completion_Rate_value1<98)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_SDCCH_Completion_Rate_BBH++;
								if(List_2G.contains(Cell_Name_Value+"SDCCH Completion Rate (BBH)"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"SDCCH Completion Rate (BBH)", SDCCH_Completion_Rate_value1);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<98");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("SDCCH Completion Rate (BBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"SDCCH Completion Rate (BBH)", SDCCH_Completion_Rate_value1);
								}	
							}
						}
						double TCH_Completion_Rate_value=0;
						if(TCH_Completion_Rate!=null)
						{  
							if(TCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								TCH_Completion_Rate_value=TCH_Completion_Rate.getNumericCellValue();
							
							}
							else if(TCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									TCH_Completion_Rate_value =Double.parseDouble(TCH_Completion_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									TCH_Completion_Rate_value = 0.0;
								}
							}
							else
							{
								TCH_Completion_Rate_value = 0.0;
							}
						}
						
						double TCH_Completion_Rate_value1=100.0-TCH_Completion_Rate_value;
						if(TCH_Completion_Rate_value1<98)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_TCH_Completion_Rate_BBH++;
								if(List_2G.contains(Cell_Name_Value+"TCH Completion Rate (BBH)"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"TCH Completion Rate (BBH)", TCH_Completion_Rate_value1);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<98");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("TCH Completion Rate (BBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"TCH Completion Rate (BBH)", TCH_Completion_Rate_value1);
									
								}
								
							}	
						}
						double RX_Quality_DL_value=0;
						if(RX_Quality_DL!=null)
						{  
							if(RX_Quality_DL.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RX_Quality_DL_value=RX_Quality_DL.getNumericCellValue();
							
							}
							else if(RX_Quality_DL.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RX_Quality_DL_value =Double.parseDouble(RX_Quality_DL.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RX_Quality_DL_value = 0.0;
								}
							}
							else
							{
								RX_Quality_DL_value = 0.0;
							}
						}
						
						if(RX_Quality_DL_value<96)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_RX_Quality_DL_BBH++;
								if(List_2G.contains(Cell_Name_Value+"RX Quality DL (0-5) (BBH)"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"RX Quality DL (0-5) (BBH)", RX_Quality_DL_value);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<96");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("RX Quality DL (0-5) (BBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"RX Quality DL (0-5) (BBH)", RX_Quality_DL_value);
								}
							
							}
						}
					}		
				}
			}
			 int Count_Handover_Success_Rate_NBH=0;
			 int Count_RX_Quality_DL_NBH=0;
			 int Count_SDCCH_Assignment_Success_NBH=0;
			 int Count_SDCCH_Completion_Rate_NBH=0;
			 int Count_TCH_Assignment_Success_NBH=0;
			 int Count_TCH_Completion_Rate_NBH=0;
			
			for(int i=0;i<NBH_2G_Sheet1.getPhysicalNumberOfRows();i++)
			{
				Row row=NBH_2G_Sheet1.getRow(i);
				if(row!=null)
				{
					if(i==7)
					{
						for(int k=0; k<row.getPhysicalNumberOfCells();k++)
						{
							Cell c2=row.getCell(k);
							if(c2!=null)
							{
								String ColHeading1=c2.getStringCellValue().trim();
								if(ColHeading1.equals("Cell Name"))
								{
									Cell_Name_Index1=k;
								}
								else if(ColHeading1.equals("SD_drop (%)"))
								{
									SDCCH_Completion_Rate_Index1=k;
								}
								else if(ColHeading1.equals("TCH_Drop %"))
								{
									TCH_Completion_Rate_Index1=k;
								}
								else if(ColHeading1.equals("VFE _HSR %"))
								{
									Handover_Success_Rate_Index1=k;
								}
								else if(ColHeading1.equals("VFE_SDCCH BLK%"))
								{
									SDCCH_Assignment_Success_Index1=k;
								}
								else if(ColHeading1.equals("TCH BLK BBH% VFE"))
								{
									TCH_Assignment_Success_Index1=k;
								}
								else if(ColHeading1.equals("DL_QL"))
								{
									DL_QL_Index1=k;
								}
							}
						}
					}
					Cell Cell_Name=null;
					Cell SDCCH_Completion_Rate=null;
					Cell TCH_Completion_Rate=null;
					Cell Handover_Success_Rate=null;
					Cell SDCCH_Assignment_Success=null;
					Cell TCH_Assignment_Success=null;
					Cell DL_QL=null;
					Cell_Name=row.getCell(Cell_Name_Index1);
					SDCCH_Completion_Rate=row.getCell(SDCCH_Completion_Rate_Index1);
					TCH_Completion_Rate=row.getCell(TCH_Completion_Rate_Index1);
					Handover_Success_Rate=row.getCell(Handover_Success_Rate_Index1);
					SDCCH_Assignment_Success=row.getCell(SDCCH_Assignment_Success_Index1);
					TCH_Assignment_Success=row.getCell(TCH_Assignment_Success_Index1);
					DL_QL=row.getCell(DL_QL_Index1);
					if(i>=8)
					{
						String Cell_Name_Value="";
						if(Cell_Name!=null)
						{
							if(Cell_Name.getCellType()==Cell.CELL_TYPE_STRING)
							{
								Cell_Name_Value=Cell_Name.getStringCellValue();
							}
							else if(Cell_Name.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Cell_Name_Value=String.valueOf((int)Cell_Name.getNumericCellValue());
							}	
						}
						
						double SDCCH_Completion_Rate_value=0;
						if(SDCCH_Completion_Rate!=null)
						{  
							if(SDCCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								SDCCH_Completion_Rate_value=SDCCH_Completion_Rate.getNumericCellValue();
							
							}
							else if(SDCCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									SDCCH_Completion_Rate_value =Double.parseDouble(SDCCH_Completion_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									SDCCH_Completion_Rate_value = 0.0;
								}
							}
							else
							{
								SDCCH_Completion_Rate_value = 0.0;
							}
						}
						
						double SDCCH_Completion_Rate_value_value1=100.0-SDCCH_Completion_Rate_value;
						if(SDCCH_Completion_Rate_value_value1<98.8)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_SDCCH_Completion_Rate_NBH++;
								if(List_2G.contains(Cell_Name_Value+"SDCCH Completion Rate (NBH)"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"SDCCH Completion Rate (NBH)", SDCCH_Completion_Rate_value_value1);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<98.8");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("SDCCH Completion Rate (NBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"SDCCH Completion Rate (NBH)", SDCCH_Completion_Rate_value_value1);
								}
							}
						}
						double TCH_Completion_Rate_value=0;
						if(TCH_Completion_Rate!=null)
						{  
							if(TCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								TCH_Completion_Rate_value=TCH_Completion_Rate.getNumericCellValue();
							
							}
							else if(TCH_Completion_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									TCH_Completion_Rate_value =Double.parseDouble(TCH_Completion_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									TCH_Completion_Rate_value = 0.0;
								}
							}
							else
							{
								TCH_Completion_Rate_value = 0.0;
							}
						}
						
						double TCH_Completion_Rate_value1=100.0-TCH_Completion_Rate_value;
						if(TCH_Completion_Rate_value1<98.5)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_TCH_Completion_Rate_NBH++;
								if(List_2G.contains(Cell_Name_Value+"TCH Completion Rate NBH)"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"TCH Completion Rate NBH)", TCH_Completion_Rate_value1);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<98.5");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("TCH Completion Rate NBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"TCH Completion Rate NBH)", TCH_Completion_Rate_value1);
								}
							}
						}
						double Handover_Success_Rate_value=0;
						if(Handover_Success_Rate!=null)
						{  
							if(Handover_Success_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Handover_Success_Rate_value=Handover_Success_Rate.getNumericCellValue();
							
							}
							else if(Handover_Success_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									Handover_Success_Rate_value =Double.parseDouble(Handover_Success_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									Handover_Success_Rate_value = 0.0;
								}
							}
							else
							{
								Handover_Success_Rate_value = 0.0;
							}
						}
						
						if(Handover_Success_Rate_value<96)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_Handover_Success_Rate_NBH++;
								if(List_2G.contains(Cell_Name_Value+"Handover Success Rate (NBH)"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"Handover Success Rate (NBH)", Handover_Success_Rate_value);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<96");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("Handover Success Rate (NBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"Handover Success Rate (NBH)", Handover_Success_Rate_value);
								}
							}	
						}
						double SDCCH_Assignment_Success_value=0;
						if(SDCCH_Assignment_Success!=null)
						{  
							if(SDCCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								SDCCH_Assignment_Success_value=SDCCH_Assignment_Success.getNumericCellValue();
							
							}
							else if(SDCCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									SDCCH_Assignment_Success_value =Double.parseDouble(SDCCH_Assignment_Success.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									SDCCH_Assignment_Success_value = 0.0;
								}
							}
							else
							{
								SDCCH_Assignment_Success_value = 0.0;
							}
						}
						
						double SDCCH_Assignment_Success_value1=100.0-SDCCH_Assignment_Success_value;
						if(SDCCH_Assignment_Success_value1<99.6)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_SDCCH_Assignment_Success_NBH++;
								if(List_2G.contains(Cell_Name_Value+"SDCCH Assignment Success (NBH)"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"SDCCH Assignment Success (NBH)", SDCCH_Assignment_Success_value1);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<99.6");
									ocell.setCellStyle(borderStyle);
								
									ocell=orow.createCell(1);
									ocell.setCellValue("SDCCH Assignment Success (NBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"SDCCH Assignment Success (NBH)", SDCCH_Assignment_Success_value1);
								}
								
							}
						}
						double TCH_Assignment_Success_value=0;
						if(TCH_Assignment_Success!=null)
						{  
							if(TCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								TCH_Assignment_Success_value=TCH_Assignment_Success.getNumericCellValue();
							
							}
							else if(TCH_Assignment_Success.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									TCH_Assignment_Success_value =Double.parseDouble(TCH_Assignment_Success.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									TCH_Assignment_Success_value = 0.0;
								}
							}
							else
							{
								TCH_Assignment_Success_value = 0.0;
							}
						}
						 
						
						double TCH_Assignment_Success_value1=100.0-TCH_Assignment_Success_value;
						if(TCH_Assignment_Success_value1<99)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_TCH_Assignment_Success_NBH++;
								if(List_2G.contains(Cell_Name_Value+"TCH Assignment Success (NBH)"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"TCH Assignment Success (NBH)", TCH_Assignment_Success_value1);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<99");
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(1);
									ocell.setCellValue("TCH Assignment Success (NBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"TCH Assignment Success (NBH)", TCH_Assignment_Success_value1);
								}
							}
						}
						double RX_Quality_DL_value=0;
						if(DL_QL!=null)
						{  
							if(DL_QL.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RX_Quality_DL_value=DL_QL.getNumericCellValue();
							
							}
							else if(DL_QL.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RX_Quality_DL_value =Double.parseDouble(DL_QL.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RX_Quality_DL_value = 0.0;
								}
							}
							else
							{
								RX_Quality_DL_value = 0.0;
							}
						}
						
						if(RX_Quality_DL_value<97)
						{
							if(Cell_Name_Value.trim().length()>0)
							{
								Count_RX_Quality_DL_NBH++;
								if(List_2G.contains(Cell_Name_Value+"RX Quality DL (0-5) (NBH)"))
								{
									Hash_Map_2G.put(Cell_Name_Value+"RX Quality DL (0-5) (NBH)", RX_Quality_DL_value);
								}
								else
								{
									orow=s1.createRow(s1.getPhysicalNumberOfRows());
									ocell=orow.createCell(0);
									ocell.setCellValue(Cell_Name_Value);
									ocell.setCellStyle(borderStyle);
									ocell=orow.createCell(2);
									ocell.setCellValue("<97");
									ocell.setCellStyle(borderStyle);
									
									ocell=orow.createCell(1);
									ocell.setCellValue("RX Quality DL (0-5) (NBH)");
									ocell.setCellStyle(borderStyle);
									Hash_Map_2G.put(Cell_Name_Value+"RX Quality DL (0-5) (NBH)", RX_Quality_DL_value);
								}
							}
						}
					}
				}							
			}
			 int Count_DCR_PS=0;
			 int Count_DCR_CS=0;
			 int Count_DSSR=0;
			 int Count_RAB_Succ_PS_Rate=0;
			 int Count_RAB_Succ_CS_Rate=0;
			 int Count_CSSR=0;
			 int Count_RRC_Succ_CS_Rate=0;
			 int Count_RRC_Succ_PS_Rate=0;
			for(int i=0;i<NBH_3G_Sheet1.getPhysicalNumberOfRows();i++)
			{
				Row row=NBH_3G_Sheet1.getRow(i);
				if(row!=null)
				{
					if(i==3)
					{
						for(int k=0; k<row.getPhysicalNumberOfCells();k++)
						{
							Cell c2=row.getCell(k);
							if(c2!=null)
							{
								String ColHeading1=c2.getStringCellValue().trim();
								if(ColHeading1.equals("UtranCell"))
								{
									UtranCell_Index=k;
								}
								else if(ColHeading1.equals("DCR_PS"))
								{
									DCR_PS_Index=k;
								}
								else if(ColHeading1.equals("DCR_CS"))
								{
									DCR_CS_Index=k;
								}
								else if(ColHeading1.equals("RRC_Succ_CS_Rate"))
								{
									RRC_Succ_CS_Rate_Index=k;
								}
								else if(ColHeading1.equals("RRC_Succ_PS_Rate"))
								{
									RRC_Succ_PS_Rate_Index=k;
								}
								else if(ColHeading1.equals("RAB_Succ_CS_Rate"))
								{
									RAB_Succ_CS_Rate_Index=k;
								}
								else if(ColHeading1.equals("RAB_Succ_PS_Rate"))
								{
									RAB_Succ_PS_Rate_Index=k;
								}
								else if(ColHeading1.equals("CSSR"))
								{
									CSSR_Index=k;
								}
								else if(ColHeading1.equals("DSSR"))
								{
									DSSR_Index=k;
								}
							}
						}
					}
					Cell UtranCell=null;
					Cell DCR_PS=null;
					Cell DCR_CS=null;
					Cell RRC_Succ_CS_Rate=null;
					Cell RRC_Succ_PS_Rate=null;
					Cell RAB_Succ_CS_Rate=null;
					Cell RAB_Succ_PS_Rate=null;
					Cell CSSR=null;
					Cell DSSR=null;
					UtranCell=row.getCell(UtranCell_Index);
					DCR_PS=row.getCell(DCR_PS_Index);
					DCR_CS=row.getCell(DCR_CS_Index);
					RRC_Succ_CS_Rate=row.getCell(RRC_Succ_CS_Rate_Index);
					RRC_Succ_PS_Rate=row.getCell(RRC_Succ_PS_Rate_Index);	
					RAB_Succ_CS_Rate=row.getCell(RAB_Succ_CS_Rate_Index);
					RAB_Succ_PS_Rate=row.getCell(RAB_Succ_PS_Rate_Index);
					CSSR=row.getCell(CSSR_Index);
					DSSR=row.getCell(DSSR_Index);
					if(i>=4)
					{
						String UtranCell_Value="";
						if(UtranCell!=null)
						{
							if(UtranCell.getCellType()==Cell.CELL_TYPE_STRING)
							{
								UtranCell_Value=UtranCell.getStringCellValue();
							}
							else if(UtranCell.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								UtranCell_Value=String.valueOf((int)UtranCell.getNumericCellValue());
							}	
						}
						double DCR_PS_value=0;
						if(DCR_PS!=null)
						{  
							if(DCR_PS.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								DCR_PS_value=DCR_PS.getNumericCellValue();
							}
							else if(DCR_PS.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									DCR_PS_value =Double.parseDouble(DCR_PS.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									DCR_PS_value = 0.0;
								}
							}
							else
							{
								DCR_PS_value = 0.0;
							}
						}
						
						if(DCR_PS_value>2)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_DCR_PS++;
								if(List_3G.contains(UtranCell_Value+"DCR_PS"))
								{
									Hash_Map_3G.put(UtranCell_Value+"DCR_PS", DCR_PS_value);
								}
								else
								{
									orow1=s2.createRow(s2.getPhysicalNumberOfRows());
									ocell1=orow1.createCell(0);
									ocell1.setCellValue(UtranCell_Value);
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(2);
									ocell1.setCellValue(">2");
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(1);
									ocell1.setCellValue("DCR_PS");
									ocell1.setCellStyle(borderStyle);
									Hash_Map_3G.put(UtranCell_Value+"DCR_PS", DCR_PS_value);
								}
							}
						}
						double DCR_CS_value=0;
						if(DCR_CS!=null)
						{  
							if(DCR_CS.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								DCR_CS_value=DCR_CS.getNumericCellValue();
							
							}
							else if(DCR_CS.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									DCR_CS_value =Double.parseDouble(DCR_CS.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									DCR_CS_value = 0.0;
								}
							}
							else
							{
								DCR_CS_value = 0.0;
							}
						}

						
						if(DCR_CS_value>2)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_DCR_CS++;
								if(List_3G.contains(UtranCell_Value+"DCR_CS"))
								{
									Hash_Map_3G.put(UtranCell_Value+"DCR_CS", DCR_CS_value);
								}
								else
								{
									orow1=s2.createRow(s2.getPhysicalNumberOfRows());
									ocell1=orow1.createCell(0);
									ocell1.setCellValue(UtranCell_Value);
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(2);
									ocell1.setCellValue(">2");
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(1);
									ocell1.setCellValue("DCR_CS");
									ocell1.setCellStyle(borderStyle);
									Hash_Map_3G.put(UtranCell_Value+"DCR_CS", DCR_CS_value);
								}
							}	
						}
						double RRC_Succ_CS_Rate_value=0;
						if(RRC_Succ_CS_Rate!=null)
						{  
							if(RRC_Succ_CS_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RRC_Succ_CS_Rate_value=RRC_Succ_CS_Rate.getNumericCellValue();
							}
							else if(RRC_Succ_CS_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RRC_Succ_CS_Rate_value =Double.parseDouble(RRC_Succ_CS_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RRC_Succ_CS_Rate_value = 0.0;
								}
							}
							else
							{
								RRC_Succ_CS_Rate_value = 0.0;
							}
						}

						
						if(RRC_Succ_CS_Rate_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_RRC_Succ_CS_Rate++;
								if(List_3G.contains(UtranCell_Value+"RRC_Succ_CS_Rate"))
								{
									Hash_Map_3G.put(UtranCell_Value+"RRC_Succ_CS_Rate", RRC_Succ_CS_Rate_value);
								}
								else
								{
									orow1=s2.createRow(s2.getPhysicalNumberOfRows());
									ocell1=orow1.createCell(0);
									ocell1.setCellValue(UtranCell_Value);
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(2);
									ocell1.setCellValue("<99");
									ocell1.setCellStyle(borderStyle);
									
									ocell1=orow1.createCell(1);
									ocell1.setCellValue("RRC_Succ_CS_Rate");
									ocell1.setCellStyle(borderStyle);
									Hash_Map_3G.put(UtranCell_Value+"RRC_Succ_CS_Rate", RRC_Succ_CS_Rate_value);
								}
								
							}
						}
						double RRC_Succ_PS_Rate_value=0;
						if(RRC_Succ_PS_Rate!=null)
						{  
							if(RRC_Succ_PS_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RRC_Succ_PS_Rate_value=RRC_Succ_PS_Rate.getNumericCellValue();
							
							}
							else if(RRC_Succ_PS_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RRC_Succ_PS_Rate_value =Double.parseDouble(RRC_Succ_PS_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RRC_Succ_PS_Rate_value = 0.0;
								}
							}
							else
							{
								RRC_Succ_PS_Rate_value = 0.0;
							}
						}
						
						
						if(RRC_Succ_PS_Rate_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_RRC_Succ_PS_Rate++;
								if(List_3G.contains(UtranCell_Value+"RRC_Succ_PS_Rate"))
								{
									Hash_Map_3G.put(UtranCell_Value+"RRC_Succ_PS_Rate", RRC_Succ_PS_Rate_value);
								}
								else
								{
									orow1=s2.createRow(s2.getPhysicalNumberOfRows());
									ocell1=orow1.createCell(0);
									ocell1.setCellValue(UtranCell_Value);
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(2);
									ocell1.setCellValue("<99");
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(1);
									ocell1.setCellValue("RRC_Succ_PS_Rate");
									ocell1.setCellStyle(borderStyle);
									Hash_Map_3G.put(UtranCell_Value+"RRC_Succ_PS_Rate", RRC_Succ_PS_Rate_value);
								}
								
							}
						}
						double RAB_Succ_CS_Rate_value=0;
						if(RAB_Succ_CS_Rate!=null)
						{  
							if(RAB_Succ_CS_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RAB_Succ_CS_Rate_value=RAB_Succ_CS_Rate.getNumericCellValue();
							
							}
							else if(RAB_Succ_CS_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RAB_Succ_CS_Rate_value =Double.parseDouble(RAB_Succ_CS_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RAB_Succ_CS_Rate_value = 0.0;
								}
							}
							else
							{
								RAB_Succ_CS_Rate_value = 0.0;
							}
						}	
						
						if(RAB_Succ_CS_Rate_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_RAB_Succ_CS_Rate++;
								if(List_3G.contains(UtranCell_Value+"RAB_Succ_CS_Rate"))
								{
									Hash_Map_3G.put(UtranCell_Value+"RAB_Succ_CS_Rate", RAB_Succ_CS_Rate_value);
								}
								else
								{
									orow1=s2.createRow(s2.getPhysicalNumberOfRows());
									ocell1=orow1.createCell(0);
									ocell1.setCellValue(UtranCell_Value);
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(2);
									ocell1.setCellValue("<99");
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(1);
									ocell1.setCellValue("RAB_Succ_CS_Rate");
									ocell1.setCellStyle(borderStyle);
									Hash_Map_3G.put(UtranCell_Value+"RAB_Succ_CS_Rate", RAB_Succ_CS_Rate_value);
								}
							}
						}
						double RAB_Succ_PS_Rate_value=0;
						if(RAB_Succ_PS_Rate!=null)
						{  
							if(RAB_Succ_PS_Rate.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								RAB_Succ_PS_Rate_value=RAB_Succ_PS_Rate.getNumericCellValue();
							}
							else if(RAB_Succ_PS_Rate.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									RAB_Succ_PS_Rate_value =Double.parseDouble(RAB_Succ_PS_Rate.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									RAB_Succ_PS_Rate_value = 0.0;
								}
							}
							else
							{
								RAB_Succ_PS_Rate_value = 0.0;
							}
						}	
						
						if(RAB_Succ_PS_Rate_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_RAB_Succ_PS_Rate++;
								if(List_3G.contains(UtranCell_Value+"RAB_Succ_PS_Rate"))
								{
									Hash_Map_3G.put(UtranCell_Value+"RAB_Succ_PS_Rate", RAB_Succ_PS_Rate_value);
								}
								else
								{
									orow1=s2.createRow(s2.getPhysicalNumberOfRows());
									ocell1=orow1.createCell(0);
									ocell1.setCellValue(UtranCell_Value);
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(2);
									ocell1.setCellValue("<99");
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(1);
									ocell1.setCellValue("RAB_Succ_PS_Rate");
									ocell1.setCellStyle(borderStyle);
									Hash_Map_3G.put(UtranCell_Value+"RAB_Succ_PS_Rate", RAB_Succ_PS_Rate_value);
								}
								
							}
						}
						double CSSR_value=0;
						if(CSSR!=null)
						{  
							if(CSSR.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								CSSR_value=CSSR.getNumericCellValue();
							
							}
							else if(CSSR.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									CSSR_value =Double.parseDouble(CSSR.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									CSSR_value = 0.0;
								}
							}
							else
							{
								CSSR_value = 0.0;
							}
						}
						if(CSSR_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_CSSR++;
								if(List_3G.contains(UtranCell_Value+"CSSR"))
								{
									Hash_Map_3G.put(UtranCell_Value+"CSSR", CSSR_value);
								}
								else
								{
									orow1=s2.createRow(s2.getPhysicalNumberOfRows());
									ocell1=orow1.createCell(0);
									ocell1.setCellValue(UtranCell_Value);
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(2);
									ocell1.setCellValue("<99");
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(1);
									ocell1.setCellValue("CSSR");
									ocell1.setCellStyle(borderStyle);
									Hash_Map_3G.put(UtranCell_Value+"CSSR", CSSR_value);
									
								}
								
							}
						}
						double DSSR_value=0;
						if(DSSR!=null)
						{  
							if(DSSR.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								DSSR_value=DSSR.getNumericCellValue();
							}
							else if(DSSR.getCellType()==Cell.CELL_TYPE_STRING)
							{
								try
								{
									DSSR_value =Double.parseDouble(DSSR.getStringCellValue());
								}
								catch(NumberFormatException e)	
								{
									DSSR_value = 0.0;
								}
							}
							else
							{
								DSSR_value = 0.0;
							}
						}	
						if(DSSR_value<99)
						{
							if(UtranCell_Value.trim().length()>0)
							{
								Count_DSSR++;
								if(List_3G.contains(UtranCell_Value+"DSSR"))
								{
									Hash_Map_3G.put(UtranCell_Value+"DSSR", DSSR_value);
								}
								else
								{
									orow1=s2.createRow(s2.getPhysicalNumberOfRows());
									ocell1=orow1.createCell(0);
									ocell1.setCellValue(UtranCell_Value);
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(2);
									ocell1.setCellValue("<99");
									ocell1.setCellStyle(borderStyle);
									ocell1=orow1.createCell(1);
									ocell1.setCellValue("DSSR");
									ocell1.setCellStyle(borderStyle);
									Hash_Map_3G.put(UtranCell_Value+"DSSR", DSSR_value);
								}
							}
						}
					}
				}					
			}
			
			orow=s1.getRow(0);
			int cIn0 = orow.getPhysicalNumberOfCells()-2;
			ocell=orow.createCell(cIn0);
			ocell.setCellValue(date3);
			ocell.setCellStyle(newstyle5);
		
			
			orow1=s2.getRow(0);
			int cIn1 = orow1.getPhysicalNumberOfCells()-2;
			ocell1=orow1.createCell(cIn1);
			ocell1.setCellValue(date3);
			ocell1.setCellStyle(newstyle5);
			
//			int value=orow.getPhysicalNumberOfCells();
//			int value1=orow1.getPhysicalNumberOfCells();
			for(int i=1;i<s1.getPhysicalNumberOfRows();i++)
			{
				Row row=s1.getRow(i);
				if(row!=null)
				{
					Cell cell_Name=row.getCell(0);
					Cell KPI_effected=row.getCell(1);
					String cell_Name_Value="";
					if(cell_Name!=null)
					{	
						if(cell_Name.getCellType()==Cell.CELL_TYPE_STRING)
						{
							cell_Name_Value=cell_Name.getStringCellValue();
						}	
					}
					String KPI_effected_Value="";
					if(KPI_effected!=null)
					{	
						if(KPI_effected.getCellType()==Cell.CELL_TYPE_STRING)
						{
							KPI_effected_Value=KPI_effected.getStringCellValue();
						}	
					}
					if(Hash_Map_2G.containsKey(cell_Name_Value+KPI_effected_Value))
					{	
						Cell cell=s1.getRow(i).createCell(cIn0);
						cell.setCellValue(Hash_Map_2G.get(cell_Name_Value+KPI_effected_Value));
						cell.setCellStyle(borderStyle);
					}
					else
					{
						Cell cell=s1.getRow(i).createCell(cIn0);
						cell.setCellValue("");
						cell.setCellStyle(borderStyle);
					}
				}
			}
			for(int i=1;i<s2.getPhysicalNumberOfRows();i++)
			{
				Row row=s2.getRow(i);
				if(row!=null)
				{
					Cell cell_Name=row.getCell(0);
					Cell KPI_effected=row.getCell(1);
					String cell_Name_Value="";
					if(cell_Name!=null)
					{	
						if(cell_Name.getCellType()==Cell.CELL_TYPE_STRING)
						{
							cell_Name_Value=cell_Name.getStringCellValue();
						}	
					}
					String KPI_effected_Value="";
					if(KPI_effected!=null)
					{	
						if(KPI_effected.getCellType()==Cell.CELL_TYPE_STRING)
						{
							KPI_effected_Value=KPI_effected.getStringCellValue();
						}	
					}
					if(Hash_Map_3G.containsKey(cell_Name_Value+KPI_effected_Value))
					{	
						Cell cell=s2.getRow(i).createCell(cIn1);
						cell.setCellValue(Hash_Map_3G.get(cell_Name_Value+KPI_effected_Value));
						cell.setCellStyle(borderStyle);
					}
					else
					{
						Cell cell=s2.getRow(i).createCell(cIn1);
						cell.setCellValue("");
						cell.setCellStyle(borderStyle);
					}
				}
			}
			orow=s1.getRow(0);
			ocell=orow.createCell(cIn0+1);
			ocell.setCellValue("Count");
			ocell.setCellStyle(newstyle5);
			
			orow=s1.getRow(0);
			ocell=orow.createCell(cIn0+2);
			ocell.setCellValue("Ranking");
			ocell.setCellStyle(newstyle5);
		
			for(int i=1;i<s1.getPhysicalNumberOfRows();i++)
			{
				Row row=s1.getRow(i);
				if(row!=null)
				{
					Cell cell=row.createCell(cIn0+1);
					cell.setCellValue("");
					cell.setCellStyle(borderStyle);
					
					Cell cell1=row.createCell(cIn0+2);
					cell1.setCellValue("");
					cell1.setCellStyle(borderStyle);
				}
			}
			
			
			int count=0;
			for(int i=1;i<s1.getPhysicalNumberOfRows();i++)
			{
				Row row=s1.getRow(i);
				if(row!=null)
				{
					count=0;
					for(int j=3;j<12;j++)
					{
						Cell rowwisedata=row.getCell(j);
						if(rowwisedata!=null)
						{
							if(rowwisedata.toString().trim().length()>0)
							{
								count++;
							}
						}
					}
					Cell cell=row.createCell(cIn0+1);
					cell.setCellValue(count);
					cell.setCellStyle(borderStyle);
					if(count==10)
					{
						Cell cell1=row.createCell(cIn0+2);
						cell1.setCellValue("1");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==9)
					{
						Cell cell1=row.createCell(cIn0+2);
						cell1.setCellValue("2");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==8)
					{
						Cell cell1=row.createCell(cIn0+2);
						cell1.setCellValue("3");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==7)
					{
						Cell cell1=row.createCell(cIn0+2);
						cell1.setCellValue("4");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==6)
					{
						Cell cell1=row.createCell(cIn0+2);
						cell1.setCellValue("5");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==5)
					{
						Cell cell1=row.createCell(cIn0+2);
						cell1.setCellValue("6");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==4)
					{
						Cell cell1=row.createCell(cIn0+2);
						cell1.setCellValue("7");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==3)
					{
						Cell cell1=row.createCell(cIn0+2);
						cell1.setCellValue("8");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==2)
					{
						Cell cell1=row.createCell(cIn0+2);
						cell1.setCellValue("9");
						cell1.setCellStyle(borderStyle);
					}
					else if(count==1)
					{
						Cell cell1=row.createCell(cIn0+2);
						cell1.setCellValue("10");
						cell1.setCellStyle(borderStyle);
					}
				}
			}
			
			orow1=s2.getRow(0);
			ocell1=orow1.createCell(cIn1+1);
			ocell1.setCellValue("Count");
			ocell1.setCellStyle(newstyle5);
			
			orow1=s2.getRow(0);
			ocell1=orow1.createCell(cIn1+2);
			ocell1.setCellValue("Ranking");
			ocell1.setCellStyle(newstyle5);
			for(int i=1;i<s2.getPhysicalNumberOfRows();i++)
			{
				Row row=s2.getRow(i);
				if(row!=null)
				{
					Cell cell=row.createCell(cIn1+1);
					cell.setCellValue("");
					cell.setCellStyle(borderStyle);
					
					Cell cell1=row.createCell(cIn1+2);
					cell1.setCellValue("");
					cell1.setCellStyle(borderStyle);
				}
			}
			int count1=0;
			
			for(int i=1;i<s2.getPhysicalNumberOfRows();i++)
			{
				Row row=s2.getRow(i);
				if(row!=null)
				{
					count1=0;
					for(int j=3;j<12;j++)
					{
						Cell rowwisedata=row.getCell(j);
						if(rowwisedata!=null)
						{
							if(rowwisedata.toString().trim().length()>0)
							{
								count1++;
							}
							
						}
					}
					Cell cell=row.createCell(cIn1+1);
					cell.setCellValue(count1);
					cell.setCellStyle(borderStyle);
					if(count1==10)
					{
						Cell cell1=row.createCell(cIn1+2);
						cell1.setCellValue("1");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==9)
					{
						Cell cell1=row.createCell(cIn1+2);
						cell1.setCellValue("2");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==8)
					{
						Cell cell1=row.createCell(cIn1+2);
						cell1.setCellValue("3");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==7)
					{
						Cell cell1=row.createCell(cIn1+2);
						cell1.setCellValue("4");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==6)
					{
						Cell cell1=row.createCell(cIn1+2);
						cell1.setCellValue("5");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==5)
					{
						Cell cell1=row.createCell(cIn1+2);
						cell1.setCellValue("6");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==4)
					{
						Cell cell1=row.createCell(cIn1+2);
						cell1.setCellValue("7");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==3)
					{
						Cell cell1=row.createCell(cIn1+2);
						cell1.setCellValue("8");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==2)
					{
						Cell cell1=row.createCell(cIn1+2);
						cell1.setCellValue("9");
						cell1.setCellStyle(borderStyle);
					}
					else if(count1==1)
					{
						Cell cell1=row.createCell(cIn1+2);
						cell1.setCellValue("10");
						cell1.setCellStyle(borderStyle);
					}
				}
			}
			//
			orow2=s3.getRow(0);
			int cIn2 = orow2.getPhysicalNumberOfCells();
			ocell2=orow2.createCell(cIn2);
			ocell2.setCellValue(date3);
			ocell2.setCellStyle(newstyle5);
			
			// Count of Cell Sheet Write Start here
			 orow2=s3.getRow(1);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_ERCS_BSS_Dash_DL_Hard_Blocking_BH_Value1);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(2);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_ERCS_BSS_Dash_EDGE_Average_DL_THruput_PER_TBF_BH_Value1);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(3);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_ERCS_BSS_Dash_GPRS_Average_DL_THruput_PER_TBF_BH_Value1);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(4);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_ERCS_BSS_NQI_TBF_Success_Rate_Dashboard_1_BH_Value1);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(5);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_Handover_Success_Rate_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(6);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_Handover_Success_Rate_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(7);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_RX_Quality_DL_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(8);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_RX_Quality_DL_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(9);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_SDCCH_Assignment_Success_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(10);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_SDCCH_Assignment_Success_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(11);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_SDCCH_Completion_Rate_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(12);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_SDCCH_Completion_Rate_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(13);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_TCH_Assignment_Success_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(14);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_TCH_Assignment_Success_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(15);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_TCH_Completion_Rate_BBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(16);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_TCH_Completion_Rate_NBH);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(17);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_DCR_PS);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(18);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_DCR_CS);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(19);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_DSSR);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(20);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_RAB_Succ_PS_Rate);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(21);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_RAB_Succ_CS_Rate);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(22);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_CSSR);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(23);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_RRC_Succ_CS_Rate);
			 ocell2.setCellStyle(borderStyle);
			 
			 orow2=s3.getRow(24);
			 ocell2=orow2.createCell(cIn2);
			 ocell2.setCellValue(Count_RRC_Succ_PS_Rate);
			 ocell2.setCellStyle(borderStyle);
			
			List_2G.clear();
			List_3G.clear();
			Hash_Map_2G.clear();
			Hash_Map_3G.clear();
			DBBH_2G1=null;
			BBH_2G1=null;
			NBH_2G1=null;
			NBH_3G1=null;
			FileOutputStream outExcel1=new FileOutputStream(new File(path+"Output Worst Cell.xlsx"));              
			wbb4.write(outExcel1); //write the output data
			wbb4.close();
			outExcel1.close();
			wbb.close();
			wbb1.close();
			wbb2.close();
			wbb3.close();
	}

}
