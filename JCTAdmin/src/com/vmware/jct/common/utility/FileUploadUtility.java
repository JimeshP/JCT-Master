package com.vmware.jct.common.utility;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Method to fetch the ONet data from excel 
 * and return in in a map
 * @param file
 * @param fileName
 * @return OnetOccupationMap
 */

public class FileUploadUtility {
	
	private static final Logger LOGGER = LoggerFactory.getLogger(FileUploadUtility.class);
	
	public Map<String, ArrayList<String>> readOnetOccupationExcel(InputStream file, String fileName) {
		LOGGER.info(">>>>>> FileUploadUtility.readOnetOccupationExcel");
		//Map<String, String> OnetOccupationMap = new TreeMap<String, String>();
		
		HashMap<String, ArrayList<String>> OnetOccupationMap = new HashMap<String, ArrayList<String>>();
//		/ArrayList<String> arraylist;
		 /*arraylist.add("Hello");
		    arraylist.add("World.");
		    hashmap.put("my key", arraylist);
		    arraylist = hashmap.get("not inserted");*/
		    
		try{
			 Workbook workbook = null;
	            if(fileName.toLowerCase().endsWith("xlsx")){
	                workbook = new XSSFWorkbook(file);
	            }else if(fileName.toLowerCase().endsWith("xls")){
	                workbook = new HSSFWorkbook(file);
	            }
	            String title = "";
                String code =  "";
	            //Get the number of sheets in the xlsx file
	            int numberOfSheets = workbook.getNumberOfSheets();
	           //loop through each of the sheets
	            for(int i=0; i < numberOfSheets; i++){	                 
	                //Get the nth sheet from the workbook
	                Sheet sheet = workbook.getSheetAt(i);	                 
	                //every sheet has rows, iterate over them
	                Iterator<Row> rowIterator = sheet.iterator();
	                while (rowIterator.hasNext()) 
	                {
	                    String name = "";
	                    String shortCode = "";
	                    String onetDesc = "";
	                   
	                    
	                    //Get the row object
	                    Row row = rowIterator.next();
	                     
	                    if(row.getRowNum()==0) { //Skipping the first row of the excel as its the Header row
	                    	//code = row.getCell(0).toString();
	                    	//title = row.getCell(1).toString();
	                    	code = "Code";
	                    	title = "Title";
	                    	continue;
	                    }
	                    
	                    //Every row has columns, get the column iterator and iterate over them
	                    Iterator<Cell> cellIterator = row.cellIterator();
	                      
	                    while (cellIterator.hasNext()) 
	                    {
	                        //Get the Cell object
	                        Cell cell = cellIterator.next();
	                         
	                        //check the cell type and process accordingly
	                        switch(cell.getCellType()){
	                        case Cell.CELL_TYPE_STRING:
	                            if(shortCode.equalsIgnoreCase("")){
	                                shortCode = cell.getStringCellValue().trim();
	                            }else if(name.equalsIgnoreCase("")){
	                                //2nd column
	                                name = cell.getStringCellValue().trim();
	                            }else if(onetDesc.equalsIgnoreCase("")){
	                                //3rd column
	                            	onetDesc = cell.getStringCellValue().trim();
	                            } else{
	                            	LOGGER.info("Random data::"+cell.getStringCellValue());
	                            }
	                            break;
	                        case Cell.CELL_TYPE_NUMERIC:	                        	
	                        	LOGGER.info("Random data::"+cell.getNumericCellValue());	                        	                        	
	                        }
	                    } //end of cell iterator
	                    /*if() {
	                    	
	                    }*/
	                    if((!code.equalsIgnoreCase("Code")) || (!title.equalsIgnoreCase("Title"))) {
	                    	//OnetOccupationMap.put("1", "1");
	                    	ArrayList<String> arraylist = new ArrayList<String>();
	                    	arraylist.add("1");
	                    	arraylist.add("1");
	                    	OnetOccupationMap.put("1", arraylist);
	                    } else if((!shortCode.equalsIgnoreCase("")) && (!name.equalsIgnoreCase(""))){
	                    	//OnetOccupationMap.put(shortCode, name);
	                    	ArrayList<String> arraylist = new ArrayList<String>();	                    	
	                    	if(!name.endsWith("") || name != "" ) {
	                    		arraylist.add(name);
	                    	} else {
	                    		arraylist.add("N/A");
	                    	} 
	                    	if(!onetDesc.endsWith("") || onetDesc != "" ) {
	                    		arraylist.add(onetDesc);
	                    	} else {
	                    		arraylist.add("N/A");
	                    	}                    	
	                    	OnetOccupationMap.put(shortCode, arraylist);
	                    }
	                } //end of rows iterator	                 	                 
	            } //end of sheets for loop
		} catch(Exception e) {
			LOGGER.info(e.getMessage());
		}
		LOGGER.info("<<<<<< FileUploadUtility.readOnetOccupationExcel");
		return OnetOccupationMap;
	}
}
