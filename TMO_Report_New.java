package sadagi.ericsson.softhuman.validation.model;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TMO_Report_New {
	
	String PATH="C:\\MY Work\\myPro\\SoftHumanActivity\\WebContent\\Excel\\TMO Report\\";
	
	public static void main(String[] args) 
	{
		new TMO_Report_New().doProcess();
	}
	
	private  LinkedHashMap<String, Integer> Hash_Map_summary=null;
	public void doProcess()
	{	
		File TMO_Daily=null;
		try
		{
			Hash_Map_summary= new LinkedHashMap<String , Integer>();
    		
			TMO_Daily=new File(PATH+"TMO Daily Trending Output.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(TMO_Daily));
			
			XSSFWorkbook dworkbook=new XSSFWorkbook();
			Sheet s1=dworkbook.createSheet("Sheet1");
			Sheet s2=dworkbook.createSheet("Summary");
			
			Cell ocell=null;
			Row orow=null;
			
			XSSFCellStyle newstyle = dworkbook.createCellStyle();
			byte[] rgb = new byte[3];
			rgb[0] = (byte) 0; // red
			rgb[1] = (byte) 32; // green
			rgb[2] = (byte) 96; // blue
			XSSFColor myColor = new XSSFColor(rgb);
			newstyle.setFillForegroundColor(myColor);
			newstyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			
			Font font = dworkbook.createFont();
			font.setBold(true);
			font.setColor(IndexedColors.WHITE.getIndex());
			
			newstyle.setFont(font); 
			newstyle.setAlignment(HorizontalAlignment.CENTER);
			newstyle.setVerticalAlignment(VerticalAlignment.CENTER);
			
			orow=s1.createRow(0);
			ocell=orow.createCell(0);
			ocell.setCellValue("L1900 Overlay Site");
			ocell.setCellStyle(newstyle);
			
			ocell=orow.createCell(1);
			ocell.setCellValue("Current date");
			ocell.setCellStyle(newstyle);
			
			ocell=orow.createCell(2);
			ocell.setCellValue("LTE AWS");
			ocell.setCellStyle(newstyle);
			
			ocell=orow.createCell(3);
			ocell.setCellValue("LTE PCS");
			ocell.setCellStyle(newstyle);
			
			ocell=orow.createCell(4);
			ocell.setCellValue("UMTS");
			ocell.setCellStyle(newstyle);
			
			ocell=orow.createCell(5);
			ocell.setCellValue("GSM");
			ocell.setCellStyle(newstyle);
			
			ocell=orow.createCell(6);
			ocell.setCellValue("Remarks");
			ocell.setCellStyle(newstyle);
			
			ocell=null;
			orow=null;
			orow=s2.createRow(0);
			ocell=orow.createCell(0);
			ocell.setCellValue("Row Labels");
			ocell.setCellStyle(newstyle);
			
			ocell=orow.createCell(1);
			ocell.setCellValue("Count of L1900 Overlay Site");
			ocell.setCellStyle(newstyle);
			 
			CellStyle borderStyle = dworkbook.createCellStyle();
			borderStyle.setBorderBottom(CellStyle.BORDER_THIN);
			borderStyle.setBorderLeft(CellStyle.BORDER_THIN);
			borderStyle.setBorderRight(CellStyle.BORDER_THIN);
			borderStyle.setBorderTop(CellStyle.BORDER_THIN);
			borderStyle.setAlignment(CellStyle.ALIGN_CENTER);
			
			Font Redfont = dworkbook.createFont();
			Redfont.setColor(IndexedColors.RED.getIndex());
			
			Font Yellowfont = dworkbook.createFont();
			Yellowfont.setColor(IndexedColors.YELLOW.getIndex());
			
			Sheet TMO_Sheet=null;
			String TMO_Site_Name="";
			Cell Column_A=null;
			Cell Column_B=null;
			Cell Column_E=null;
			String Column_A_Value="";
			String Column_B_Value="";
			String ColumnA_value1="";
			String date="";
			Date startDatevalue=null;
			boolean var=false;
			boolean var2=false;
			for (int i = 0; i < wb.getNumberOfSheets(); i++)
			{
				TMO_Sheet = wb.getSheetAt(i);
				TMO_Site_Name=TMO_Sheet.getSheetName().split("-")[0].trim();	
				date="";
				if(TMO_Sheet.getRow(1).getCell(4).getCellType()==Cell.CELL_TYPE_NUMERIC)
				{
					DateFormat dtt=new SimpleDateFormat("MM/dd/yyyy");
					startDatevalue=TMO_Sheet.getRow(1).getCell(4).getDateCellValue();
					date=dtt.format(startDatevalue);
				}
				orow=s1.createRow(s1.getPhysicalNumberOfRows());
				ocell=orow.createCell(0);
				ocell.setCellValue(TMO_Site_Name);
				ocell.setCellStyle(borderStyle);
				
				ocell=orow.createCell(1);
				ocell.setCellValue(date);
				ocell.setCellStyle(borderStyle);
				
				StringBuffer SS= new StringBuffer();
				StringBuffer SS1= new StringBuffer();
				StringBuffer SS2= new StringBuffer();
				StringBuffer SS3= new StringBuffer();
				StringBuffer SS4= new StringBuffer();
				StringBuffer SS5= new StringBuffer();
				StringBuffer SS6= new StringBuffer();
				StringBuffer SS7= new StringBuffer();
				
				var=false;
				var2=false;
				
				for(int j=1;j<TMO_Sheet.getLastRowNum()+1;j++)
				{
					Row row=TMO_Sheet.getRow(j);
					if(row!=null)
					{
						Column_A=row.getCell(0);
						Column_B=row.getCell(1);
						Column_E=row.getCell(4);
						Column_A_Value="";
						if(Column_A!=null)
						{
							if(Column_A.getCellType()==Cell.CELL_TYPE_STRING)
							{
								Column_A_Value=Column_A.getStringCellValue();
							}
							else if(Column_A.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Column_A_Value=String.valueOf(Column_A.getNumericCellValue());
							}
						}
						Column_B_Value="";
						if(Column_B!=null)
						{
							if(Column_B.getCellType()==Cell.CELL_TYPE_STRING)
							{
								Column_B_Value=Column_B.getStringCellValue();
							}
							else if(Column_B.getCellType()==Cell.CELL_TYPE_NUMERIC)
							{
								Column_B_Value=String.valueOf(Column_B.getNumericCellValue());
							}
						}
						if(Column_E!=null)
						{
							if(Column_A_Value.equals(""))
							{
								Column_A_Value=ColumnA_value1;
							}
							else
							{
								ColumnA_value1=Column_A_Value;
							}
							System.out.println(j+"=="+Column_A_Value+"=="+Column_B_Value);
							if(Column_E.getCellStyle().getFillForegroundColorColor() != null)
							{
								XSSFCellStyle cs = (XSSFCellStyle) Column_E.getCellStyle(); 
								if(cs.getFillForegroundColorColor().getARGBHex().equals("FFFF3300"))
								{
									var=true;
									if(Column_B_Value.equals("LTE AWS"))
									{
										SS.append(","+Column_A_Value);
									}
									else if(Column_B_Value.equals("LTE PCS"))
									{
										SS1.append(","+Column_A_Value);
									}
									else if(Column_B_Value.contains("UMTS"))
									{
										SS2.append(","+Column_A_Value);
									}
									else if(Column_B_Value.equals("GSM"))
									{
										SS3.append(","+Column_A_Value);
									}
								}
								else if(cs.getFillForegroundColorColor().getARGBHex().equals("FFFF9800"))
								{
									var2=true;
									if(Column_B_Value.equals("LTE AWS"))
									{
										SS4.append(","+Column_A_Value);
									}
									else if(Column_B_Value.equals("LTE PCS"))
									{
										SS5.append(","+Column_A_Value);
									}
									else if(Column_B_Value.contains("UMTS"))
									{
										SS6.append(","+Column_A_Value);
									}
									else if(Column_B_Value.equals("GSM"))
									{
										SS7.append(","+Column_A_Value);
									}
								}
									
							}
						}
					}
				}
				if((SS.length()>0) && (SS4.length()>0))
				{	
					XSSFRichTextString richString = new XSSFRichTextString(SS.toString().substring(1)+","+SS4.toString().substring(1));
					ocell=orow.createCell(2);
					richString.applyFont(0, SS.toString().substring(1).length(),Redfont);
					
					richString.applyFont(SS.toString().substring(1).length()+1, SS.toString().substring(1).length()+1+SS4.toString().substring(1).length(),Yellowfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
					
				}
				if((SS.length()>0) && (SS4.length()==0))
				{
					XSSFRichTextString richString = new XSSFRichTextString(SS.toString().substring(1));
					ocell=orow.createCell(2);
					richString.applyFont(0, SS.toString().substring(1).length(),Redfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
				}
				if((SS.length()==0) && (SS4.length()>0))
				{
					XSSFRichTextString richString = new XSSFRichTextString(SS4.toString().substring(1));
					ocell=orow.createCell(2);
					richString.applyFont(0, SS4.toString().substring(1).length(),Redfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
				}
				////
				if((SS1.length()>0)  && (SS5.length()>0))
				{	
					XSSFRichTextString richString = new XSSFRichTextString(SS1.toString().substring(1)+","+SS5.toString().substring(1));
					ocell=orow.createCell(3);
					richString.applyFont(0, SS1.toString().substring(1).length(),Redfont);
					
					richString.applyFont(SS1.toString().substring(1).length()+1, SS1.toString().substring(1).length()+1+SS5.toString().substring(1).length(),Yellowfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
				}
				if((SS1.length()>0) && (SS5.length()==0))
				{
					XSSFRichTextString richString = new XSSFRichTextString(SS1.toString().substring(1));
					ocell=orow.createCell(3);
					richString.applyFont(0, SS1.toString().substring(1).length(),Redfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
				}
				if((SS1.length()==0) && (SS5.length()>0))
				{
					XSSFRichTextString richString = new XSSFRichTextString(SS5.toString().substring(1));
					ocell=orow.createCell(3);
					richString.applyFont(0, SS5.toString().substring(1).length(),Redfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
				}
				///////
				
				if((SS2.length()>0)  && (SS6.length()>0))
				{	
					XSSFRichTextString richString = new XSSFRichTextString(SS2.toString().substring(1)+","+SS6.toString().substring(1));
					ocell=orow.createCell(4);
					richString.applyFont(0, SS2.toString().substring(1).length(),Redfont);
					
					richString.applyFont(SS2.toString().substring(1).length()+1, SS2.toString().substring(1).length()+1+SS6.toString().substring(1).length(),Yellowfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
				}
				if((SS2.length()>0) && (SS6.length()==0))
				{
					XSSFRichTextString richString = new XSSFRichTextString(SS2.toString().substring(1));
					ocell=orow.createCell(4);
					richString.applyFont(0, SS2.toString().substring(1).length(),Redfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
				}
				
				if((SS2.length()==0) && (SS6.length()>0))
				{
					XSSFRichTextString richString = new XSSFRichTextString(SS6.toString().substring(1));
					ocell=orow.createCell(4);
					richString.applyFont(0, SS6.toString().substring(1).length(),Redfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
				}
				///
				
				if((SS3.length()>0)  && (SS7.length()>0))
				{	
					XSSFRichTextString richString = new XSSFRichTextString(SS3.toString().substring(1)+","+SS7.toString().substring(1));
					ocell=orow.createCell(5);
					richString.applyFont(0, SS3.toString().substring(1).length(),Redfont);
					
					richString.applyFont(SS3.toString().substring(1).length()+1, SS3.toString().substring(1).length()+1+SS7.toString().substring(1).length(),Yellowfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
				}
				if((SS3.length()>0) && (SS7.length()==0))
				{
					XSSFRichTextString richString = new XSSFRichTextString(SS3.toString().substring(1));
					ocell=orow.createCell(5);
					richString.applyFont(0, SS3.toString().substring(1).length(),Redfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
				}
				if((SS3.length()==0) && (SS7.length()>0))
				{
					XSSFRichTextString richString = new XSSFRichTextString(SS7.toString().substring(1));
					ocell=orow.createCell(5);
					richString.applyFont(0, SS7.toString().substring(1).length(),Redfont);
					ocell.setCellValue(richString);
					ocell.setCellStyle(borderStyle);
				}
				////
				if(var==true)
				{
					ocell=orow.createCell(6);
					ocell.setCellValue("RED KPI");
					ocell.setCellStyle(borderStyle);
				}
				if((var==false) && (var2==true))
				{
					ocell=orow.createCell(6);
					ocell.setCellValue("Yellow KPI");
					ocell.setCellStyle(borderStyle);
				}
				if((var==false) && (var2==false))
				{
					ocell=orow.createCell(6);
					ocell.setCellValue("All KPIs are good");
					ocell.setCellStyle(borderStyle);
				}
				
			}
			
			
			for(int i=1;i<s1.getPhysicalNumberOfRows();i++)
			{
				Row row=s1.getRow(i);
				if(row!=null)
				{
					for(int j=0;j<7;j++)
					{
						Cell cell=row.getCell(j);
						if(cell==null)
						{
							orow=s1.getRow(i);
							ocell=orow.createCell(j);
							ocell.setCellValue("");
							ocell.setCellStyle(borderStyle);
						}
					}
						
					Cell Remark=row.getCell(6);
					if(Remark!=null)
					{
						String Remarkvalue=Remark.getStringCellValue();
						if(Hash_Map_summary.containsKey(Remarkvalue)==false)
						{
							Hash_Map_summary.put(Remarkvalue, 1);
						}
						else
						{
							Hash_Map_summary.put(Remarkvalue, Hash_Map_summary.get(Remarkvalue)+1);

						}
					}
				}
			}
			int total=0;
			for(Map.Entry<String, Integer> EN : Hash_Map_summary.entrySet())
			{	
				orow=s2.createRow(s2.getPhysicalNumberOfRows());
				ocell=orow.createCell(0);
				ocell.setCellValue(EN.getKey());
				ocell.setCellStyle(borderStyle);
				
				ocell=orow.createCell(1);
				ocell.setCellValue(EN.getValue());
				ocell.setCellStyle(borderStyle);
				
				total=total+EN.getValue();
			}
			orow=s2.createRow(s2.getPhysicalNumberOfRows());
			ocell=orow.createCell(0);
			ocell.setCellValue("Grand Total");
			ocell.setCellStyle(newstyle);
			
			ocell=orow.createCell(1);
			ocell.setCellValue(total);
			ocell.setCellStyle(newstyle);
			
			Hash_Map_summary.clear();
			System.gc();
			System.runFinalization();
			
			FileOutputStream outExcel=new FileOutputStream("C:\\MY Work\\myPro\\SoftHumanActivity\\WebContent\\Excel\\TMO Report\\Output_TMO_Report.xlsx");           
			dworkbook.write(outExcel); //write the output data
			outExcel.close(); 
			dworkbook.close();
			wb.close();
			
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}	
	}
}
