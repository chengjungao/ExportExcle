package com.dinfo.cjg;
import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;



public   class  ExcelTools {
	public static  void readxml(ArrayList<List<String>> list,String xml){
		SAXReader saxReader=new SAXReader();
		try {
			Document document = saxReader.read(new File(xml));
			Iterator iterator=document.selectNodes("/bb/bb1").iterator();
			if(iterator.hasNext()) {
				Element element = (Element) iterator.next();
				Iterator iter=element.elementIterator();
				while (iter.hasNext()) {
					//每行
					Element elem = (Element) iter.next();
					String [] colName=elem.getText().split(",");
					List<String> list1=Arrays.asList(colName);
					list.add(list1);
				}
			}
		} catch (DocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	/**
	 * 
	 * @param list
	 * @return
	 */
	public static HSSFWorkbook downWorkbook(ArrayList<List<String>> list){
		 HSSFWorkbook book =new HSSFWorkbook();
		  HSSFCellStyle style = book.createCellStyle(); // 样式对象
		  style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直   
		  style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//
		  style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
	      style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
	      style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
	      style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
	      style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
		  //style.setWrapText(true);
		  HSSFFont font1 = book.createFont();
	      font1.setFontHeightInPoints((short) 11);//字号
	      style.setFont(font1);
	      HSSFSheet sheet = book.createSheet("1");
	      sheet.autoSizeColumn(( short ) 0 );
		  sheet.setDefaultRowHeightInPoints(100);
		  sheet.setDefaultColumnWidth(15);
		  HSSFCell cell=null;
		  CellRangeAddress rangeAddress=null;
		  Map<Integer, Integer> columnMap=new HashMap<Integer, Integer>();
		  for (int i = 0; i < list.size(); i++) {
			 List<String> rowList=list.get(i);
			 HSSFRow row=sheet.getRow(i);
			 if (row==null) {
				row =sheet.createRow(i);
			}
				for (int j = 0,y=0,k_y=0;j < rowList.size(); j++,y++) {
					if(rowList.get(j)!=null&&rowList.get(j).equals("-100")){ //如果是-100
						k_y++;  
						//如果-100结束
						if(rowList.get(j+1)==null||!rowList.get(j+1).equals("-100")){
							rangeAddress=new CellRangeAddress(i,i,y,y+k_y);//合并单元格  
							sheet.addMergedRegion(rangeAddress);
						}
						y--;  //如果是  -100   y要减1
						continue;
					}
					if(rowList.get(j)!=null&&rowList.get(j).equals("-200")){
						if(!columnMap.containsKey(j)){
							columnMap.put(j,i);//将要上下合并的单元格 记录下来   <列,行>
						}
						//下一行不为-200，即不用合并。
						if(!list.get(i).get(j).equals("-200")){
							rangeAddress=new CellRangeAddress(columnMap.get(j),i+1,j,j);
							HSSFRow row2=sheet.getRow(i);
							if(row2==null){
								row2=sheet.createRow(i);
							}
							HSSFCell cell2=row2.createCell(j);
							cell2.setCellStyle(style);
							sheet.addMergedRegion(rangeAddress);
						}
						continue;
					}
					if(columnMap.containsKey(j)){
						cell=sheet.getRow(columnMap.get(j)).createCell(j);
						String value=list.get(i-1).get(j);
						//如果是8位以内的数字，就将其字符串转换成double型。没有什么特别的意义，只是为了导出成excel时，数字前不出现单引号，样式好看点。
						if(value!=null && value.matches("\\-?\\d{0,6}\\.?\\d{0,2}")&&!value.equals("0")){
							cell.setCellValue(Double.parseDouble(value));
						}else if(value!=null&&!value.equals("0")){
							cell.setCellValue(value);
						}
						cell.setCellStyle(style);
						//合并完成;
						columnMap.remove(j);
						continue;
					}
					cell=(k_y!=0)?row.createCell(y):row.createCell(j);
					String value=rowList.get(j);
					if(value!=null && value.matches("\\-?\\d{0,6}\\.?\\d{0,2}")&&!value.equals("0")){
						cell.setCellValue(Double.parseDouble(value));
					}else if(value!=null&&!value.equals("0")){
						cell.setCellValue(value);
					}
					cell.setCellStyle(style);
					//合并完成，将k_y设置为0;
					k_y=0;y=j;
				}
			 
		}
		  return book;
	}
	public static  HSSFWorkbook writeWorkbook(ArrayList<area_moneyDto> list,int key){
		  HSSFWorkbook book =new HSSFWorkbook();
		  /***************设置样式**********/
		  HSSFCellStyle style = book.createCellStyle(); // 样式对象
		  style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直   
		  style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//
		  style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
	      style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
	      style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
	      style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
	      style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
		  //style.setWrapText(true);
		  HSSFFont font1 = book.createFont();
	      font1.setFontHeightInPoints((short) 11);//字号
	      style.setFont(font1);
	      HSSFSheet sheet = book.createSheet("1");
	      sheet.autoSizeColumn(( short ) 0 );
		  sheet.setDefaultRowHeightInPoints(100);
		  sheet.setDefaultColumnWidth(15);
		 // String [] year={"2010","2011","2012","2013","2014","2015"};
		  //final String [] keyword={"证券纠纷","基金纠纷","期货纠纷"};
		  String [] rows={"安徽","北京","大连","福建","甘肃","广东","广西","贵州","海南","河北","河南","黑龙江","湖北","湖南","吉林","江苏","江西","辽宁","内蒙","宁波","宁夏","青岛","青岛","青海","厦门","山东","山西","陕西","上海","深圳","四川","天津","西藏","新疆","云南","浙江","重庆"};
		  for (int i = 0; i < rows.length; i++) {
			 HSSFRow row=sheet.createRow(i);
			  HSSFCell cell=row.createCell(0);
			  cell.setCellStyle(style);
			  cell.setCellValue(rows[i]); 
			  for (int j = 1; j < 73; j++) {
				  HSSFCell cell0=row.createCell(j);
				  cell0.setCellStyle(style);
				  cell0.setCellValue(0);
			}
			 for (int j = 0; j < list.size(); j++) {
				 area_moneyDto dto =list.get(j);
				 if(dto.getArea().equals(rows[i]))//如果地区相同
				 {
					 int index=1;//单元格   索引
					 switch (dto.getTime()) {
					 case "2010":index=1;break;
					 case "2011":index=13;break;
					 case "2012":index=25;break;
					 case "2013":index=37;break;
					 case "2014":index=49;break;
					 case "2015":index=61;break;
					}
					 switch (key) {
					case 1:
						switch (dto.getKeyword()) {
						case "证券":;break;
						case "基金":index+=4;break;
						case "期货":index+=8;break;
						}
						break;
					case 2:
						switch (dto.getKeyword()) {
						case "虚假陈述":;break;
						case "内幕交易":index+=4;break;
						case "关联交易":index+=8;break;
						}
						break;
					case 3:
						switch (dto.getKeyword()) {
						case "证券公司":;break;
						case "基金公司":index+=4;break;
						case "期货公司":index+=8;break;
						}
						break;
					default:
						break;
					}
					  HSSFCell cell1=row.getCell(index); 
				      HSSFCell cell2=row.getCell(index+1);
				      HSSFCell cell3=row.getCell(index+2);
				      HSSFCell cell4=row.getCell(index+3);
				      cell1.setCellValue(Integer.parseInt(dto.getCaseNum()));
				      cell2.setCellValue(Integer.parseInt(dto.getPeopleNum()));
				      cell3.setCellValue(Double.parseDouble(dto.getReq_amount())<0?-Double.parseDouble(dto.getReq_amount()):Double.parseDouble(dto.getReq_amount()));
				      cell4.setCellValue(Double.parseDouble(dto.getDeter_amount())<0?-Double.parseDouble(dto.getDeter_amount()):Double.parseDouble(dto.getDeter_amount()));
				      list.remove(dto);
				      continue;
				 }
			}
		}
	  return book;
	}
}
