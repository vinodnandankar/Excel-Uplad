package com.example.demo.service;

import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.model.DataSheet;
@Service
public class UploadSheetService {
	List<DataSheet> dataSheetList= new ArrayList<DataSheet>();	
		public List<DataSheet> upload(MultipartFile file)throws Exception{
			
			Workbook workbook=getWorkBook(file);
			Sheet sheet= workbook.getSheetAt(0);
			Iterator<Row> rows=sheet.iterator();
			rows.next();
			while (rows.hasNext()) {
				DataSheet dataSheet=new DataSheet();
			Row row=rows.next();
			
			dataSheet.setSrNo(getIntegerValue(row.getCell(0)));
			dataSheet.setApiVersion(getStringValue(row.getCell(1)));
			dataSheet.setApiName(getStringValue(row.getCell(2)));
			dataSheet.setApiType(getStringValue(row.getCell(3)));
			dataSheet.setApiRiskClassificatin(getStringValue(row.getCell(4)));
			dataSheet.setRamlReviewStatus(getStringValue(row.getCell(5)));
			dataSheet.setRamlReviewDate(getDateValue(row.getCell(6)));
			dataSheet.setVeracodeStatus(getStringValue(row.getCell(7)));
			dataSheet.setVeracodeDate(getDateValue(row.getCell(8)));
			dataSheet.setPenTestStatus(getStringValue(row.getCell(9)));
			dataSheet.setPenTestDate(getDateValue(row.getCell(10)));
			dataSheet.setVeracodeSlaBreach(getStringValue(row.getCell(11)));
			dataSheet.setPenTestSlaBreach(getStringValue(row.getCell(12)));
			dataSheet.setRamlReviewPending(getStringValue(row.getCell(13)));
			//dataSheet.setRiskScore(getIntegerValue(row.getCell(14)));
			dataSheet.setRiskScore(getRiskScore(dataSheet));

			dataSheetList.add(dataSheet);
			}
			return dataSheetList;
			
}
		private Workbook getWorkBook(MultipartFile file) {
			Workbook workBook=null;
			String extension=FilenameUtils.getExtension(file.getOriginalFilename());
			try {
				if(extension.equalsIgnoreCase("xlsx")) {
					workBook=new XSSFWorkbook(file.getInputStream());
				}else if(extension.equalsIgnoreCase("xls")) {
					workBook=new HSSFWorkbook(file.getInputStream());
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			return workBook;
		}
		
		private String getStringValue(Cell cell) {
			return (cell != null && cell.getCellType() == CellType.STRING) ?  cell.getStringCellValue() : "";
		}

		private Date getDateValue(Cell cell) {
			return cell != null ?  cell.getDateCellValue() : null;
		}

		private Integer getIntegerValue(Cell cell) {
			return (cell != null && cell.getCellType() == CellType.NUMERIC) ?  (int)cell.getNumericCellValue() : null;
		}

		private Boolean getBooleanValue(Cell cell) {
			return (cell != null && cell.getCellType() == CellType.BOOLEAN) ?  cell.getBooleanCellValue() : Boolean.FALSE;
		}
		
		private int getRiskScore(DataSheet dataSheet) {
			int riskRiskScore = 0;
			
			return riskRiskScore;
		}
		
}
/*Cell cell = row.getCell(6);
System.out.println("Cell================="+cell);
if (cell == null) {
      // This cell is empty/blank/un-used, handle as needed
	dataSheet.setRamlReviewDate(null);
    } else {
    	dataSheet.setRamlReviewDate(row.getCell(6).getDateCellValue());
//          String cellStr = fmt.formatCell(cell);
       // Do something with the value
    }*/