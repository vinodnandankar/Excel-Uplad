package com.example.demo.service;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
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
			
			dataSheet.setSrNo((int)row.getCell(0).getNumericCellValue());
			dataSheet.setApiVersion(row.getCell(1).getStringCellValue());
			dataSheet.setApiName(row.getCell(2).getStringCellValue());
			dataSheet.setApiType(row.getCell(3).getStringCellValue());
			dataSheet.setApiRiskClassificatin(row.getCell(4).getStringCellValue());
			dataSheet.setRamlReviewStatus(row.getCell(5).getStringCellValue());
			Cell cellRamlReviewDate = row.getCell(6);
			System.out.println("cellRamlReviewDate================="+cellRamlReviewDate);
			if (cellRamlReviewDate == null) {
				dataSheet.setRamlReviewDate(null);
			    } else {
			    	dataSheet.setRamlReviewDate(row.getCell(6).getDateCellValue());
			    }
			dataSheet.setVeracodeStatus(row.getCell(7).getStringCellValue());
			Cell cellVeracodeDate = row.getCell(8);
			System.out.println("cellVeracodeDate================="+cellVeracodeDate);
			if (cellVeracodeDate == null) {
				dataSheet.setVeracodeDate(null);
			    } else {
					dataSheet.setVeracodeDate(row.getCell(8).getDateCellValue());
			    }
			dataSheet.setPenTestStatus(row.getCell(9).getStringCellValue());
			Cell cellPenTestDate = row.getCell(10);
			System.out.println("cellPenTestDate================="+cellPenTestDate);
			if (cellPenTestDate == null) {
				dataSheet.setPenTestDate(null);
			    } else {
					dataSheet.setPenTestDate(row.getCell(10).getDateCellValue());
			    }
			Cell cellVeracodeSlaBreach = row.getCell(11);
			System.out.println("cellVeracodeSlaBreach================="+cellVeracodeSlaBreach);
			if (cellVeracodeSlaBreach == null) {
				dataSheet.setVeracodeSlaBreach(null);
			    } else {
					dataSheet.setVeracodeSlaBreach(row.getCell(11).getStringCellValue());
			    }
			Cell cellPenTestSlaBreach = row.getCell(12);
			System.out.println("cellPenTestSlaBreach================="+cellPenTestSlaBreach);
			if (cellPenTestSlaBreach == null) {
				dataSheet.setPenTestSlaBreach(null);
			    } else {
					dataSheet.setPenTestSlaBreach(row.getCell(12).getStringCellValue());
			    }
			dataSheet.setRamlReviewPending(row.getCell(13).getStringCellValue());
			dataSheet.setRiskScore((int)row.getCell(14).getNumericCellValue());

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