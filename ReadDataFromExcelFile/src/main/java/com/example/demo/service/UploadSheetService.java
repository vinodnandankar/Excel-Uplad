package com.example.demo.service;

import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

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
import com.example.demo.util.ConstantsVars;

@Service
public class UploadSheetService {
	List<DataSheet> dataSheetList = new ArrayList<DataSheet>();

	public List<DataSheet> upload(MultipartFile file) throws Exception {

		Workbook workbook = getWorkBook(file);
		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rows = sheet.iterator();
		rows.next();
		while (rows.hasNext()) {
			DataSheet dataSheet = new DataSheet();
			Row row = rows.next();

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
			// dataSheet.setRiskScore(getIntegerValue(row.getCell(14)));
			dataSheet.setRiskScore(getRiskScore(dataSheet));
			dataSheet.setOverallRiskClassification(getOverallRiskClassification(dataSheet.getRiskScore()));

			dataSheetList.add(dataSheet);
		}
		return dataSheetList;

	}

	private Workbook getWorkBook(MultipartFile file) {
		Workbook workBook = null;
		String extension = FilenameUtils.getExtension(file.getOriginalFilename());
		try {
			if (extension.equalsIgnoreCase("xlsx")) {
				workBook = new XSSFWorkbook(file.getInputStream());
			} else if (extension.equalsIgnoreCase("xls")) {
				workBook = new HSSFWorkbook(file.getInputStream());
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return workBook;
	}

	private String getStringValue(Cell cell) {
		return (cell != null && cell.getCellType() == CellType.STRING) ? cell.getStringCellValue() : "";
	}

	private Date getDateValue(Cell cell) {
		return cell != null ? cell.getDateCellValue() : null;
	}

	private Integer getIntegerValue(Cell cell) {
		return (cell != null && cell.getCellType() == CellType.NUMERIC) ? (int) cell.getNumericCellValue() : null;
	}

	private Boolean getBooleanValue(Cell cell) {
		return (cell != null && cell.getCellType() == CellType.BOOLEAN) ? cell.getBooleanCellValue() : Boolean.FALSE;
	}

	private int getRiskScore(DataSheet dataSheet) {

		Predicate<DataSheet> pendingRamlReviewStatusPredicate = api -> "Pending"
				.equalsIgnoreCase(api.getRamlReviewStatus());
		Predicate<DataSheet> pendingVeracodeStatusPredicate = api -> "Pending"
				.equalsIgnoreCase(api.getVeracodeStatus());
		Predicate<DataSheet> pendingPenTestStatusPredicate = api -> "Pending".equalsIgnoreCase(api.getPenTestStatus());
		Predicate<DataSheet> veracodeSlaBreachPredicate = api -> "SLA Breached"
				.equalsIgnoreCase(api.getVeracodeSlaBreach());
		Predicate<DataSheet> penTestSlaBreachPredicate = api -> "SLA Breached"
				.equalsIgnoreCase(api.getPenTestSlaBreach());

		Predicate<DataSheet> externalAPITypePredicate = api -> "External".equalsIgnoreCase(api.getApiType());
		Predicate<DataSheet> internalAPITypePredicate = api -> "Internal".equalsIgnoreCase(api.getApiType());

		Predicate<DataSheet> crticalRiskClassificationPredicate = api -> "Critical"
				.equalsIgnoreCase(api.getApiRiskClassificatin());
		Predicate<DataSheet> highRiskClassificationPredicate = api -> "High"
				.equalsIgnoreCase(api.getApiRiskClassificatin());
		Predicate<DataSheet> mediumRiskClassificationPredicate = api -> "Medium"
				.equalsIgnoreCase(api.getApiRiskClassificatin());
		Predicate<DataSheet> lowRiskClassificationPredicate = api -> "Low"
				.equalsIgnoreCase(api.getApiRiskClassificatin());
		int riskScore = 0;
		int tempRiskScore = 0;
		if ("Internal".equals(dataSheet.getApiType())) {
			tempRiskScore = 0;
			if ("Low".equalsIgnoreCase(dataSheet.getApiRiskClassificatin())) {

				if (pendingPenTestStatusPredicate.equals(dataSheet.getPenTestStatus())) {
					tempRiskScore = tempRiskScore + 0;
				}
				if (ConstantsVars.PENDING.equalsIgnoreCase(dataSheet.getRamlReviewStatus())) {
					tempRiskScore = tempRiskScore + 0;
				}
				if (ConstantsVars.PENDING.equalsIgnoreCase(dataSheet.getVeracodeStatus())) {
					tempRiskScore = tempRiskScore + 1;
				}
				if (ConstantsVars.SLA_BREACH.equalsIgnoreCase(dataSheet.getPenTestSlaBreach())) {
					tempRiskScore = tempRiskScore + 0;
				}
				if (ConstantsVars.SLA_BREACH.equalsIgnoreCase(dataSheet.getVeracodeSlaBreach())) {
					tempRiskScore = tempRiskScore + 0;
				}
			} else if ("Medium".equalsIgnoreCase(dataSheet.getApiRiskClassificatin())) {
				if ("Pending".equalsIgnoreCase(dataSheet.getPenTestStatus())) {
					tempRiskScore = tempRiskScore + 0;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getRamlReviewStatus())) {
					tempRiskScore = tempRiskScore + 0;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getVeracodeStatus())) {
					tempRiskScore = tempRiskScore + 2;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getPenTestSlaBreach())) {
					tempRiskScore = tempRiskScore + 0;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getVeracodeSlaBreach())) {
					tempRiskScore = tempRiskScore + 1;
				}
			} else if ("High".equalsIgnoreCase(dataSheet.getApiRiskClassificatin())) {
				if ("Pending".equalsIgnoreCase(dataSheet.getPenTestStatus())) {
					tempRiskScore = tempRiskScore + 6;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getRamlReviewStatus())) {
					tempRiskScore = tempRiskScore + 4;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getVeracodeStatus())) {
					tempRiskScore = tempRiskScore + 6;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getPenTestSlaBreach())) {
					tempRiskScore = tempRiskScore + 4;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getVeracodeSlaBreach())) {
					tempRiskScore = tempRiskScore + 4;
				}
			} else if ("Critical".equalsIgnoreCase(dataSheet.getApiRiskClassificatin())) {
				if ("Pending".equalsIgnoreCase(dataSheet.getPenTestStatus())) {
					tempRiskScore = tempRiskScore + 10;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getRamlReviewStatus())) {
					tempRiskScore = tempRiskScore + 8;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getVeracodeStatus())) {
					tempRiskScore = tempRiskScore + 10;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getPenTestSlaBreach())) {
					tempRiskScore = tempRiskScore + 8;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getVeracodeSlaBreach())) {
					tempRiskScore = tempRiskScore + 8;
				}
			}

		} else if ("External".equals(dataSheet.getApiType())) {
			tempRiskScore = 0;
			if ("Low".equalsIgnoreCase(dataSheet.getApiRiskClassificatin())) {
				if ("Pending".equalsIgnoreCase(dataSheet.getPenTestStatus())) {
					tempRiskScore = tempRiskScore + 2;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getRamlReviewStatus())) {
					tempRiskScore = tempRiskScore + 4;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getVeracodeStatus())) {
					tempRiskScore = tempRiskScore + 2;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getPenTestSlaBreach())) {
					tempRiskScore = tempRiskScore + 1;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getVeracodeSlaBreach())) {
					tempRiskScore = tempRiskScore + 1;
				}
			} else if ("Medium".equalsIgnoreCase(dataSheet.getApiRiskClassificatin())) {
				if ("Pending".equalsIgnoreCase(dataSheet.getPenTestStatus())) {
					tempRiskScore = tempRiskScore + 4;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getRamlReviewStatus())) {
					tempRiskScore = tempRiskScore + 0;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getVeracodeStatus())) {
					tempRiskScore = tempRiskScore + 4;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getPenTestSlaBreach())) {
					tempRiskScore = tempRiskScore + 2;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getVeracodeSlaBreach())) {
					tempRiskScore = tempRiskScore + 2;
				}
			} else if ("High".equalsIgnoreCase(dataSheet.getApiRiskClassificatin())) {
				if ("Pending".equalsIgnoreCase(dataSheet.getPenTestStatus())) {
					tempRiskScore = tempRiskScore + 12;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getRamlReviewStatus())) {
					tempRiskScore = tempRiskScore + 4;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getVeracodeStatus())) {
					tempRiskScore = tempRiskScore + 12;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getPenTestSlaBreach())) {
					tempRiskScore = tempRiskScore + 6;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getVeracodeSlaBreach())) {
					tempRiskScore = tempRiskScore + 6;
				}
			} else if ("Critical".equalsIgnoreCase(dataSheet.getApiRiskClassificatin())) {
				if ("Pending".equalsIgnoreCase(dataSheet.getPenTestStatus())) {
					tempRiskScore = tempRiskScore + 16;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getRamlReviewStatus())) {
					tempRiskScore = tempRiskScore + 8;
				}
				if ("Pending".equalsIgnoreCase(dataSheet.getVeracodeStatus())) {
					tempRiskScore = tempRiskScore + 16;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getPenTestSlaBreach())) {
					tempRiskScore = tempRiskScore + 10;
				}
				if ("SLA Breached".equalsIgnoreCase(dataSheet.getVeracodeSlaBreach())) {
					tempRiskScore = tempRiskScore + 10;
				}
			}
		}
		riskScore = tempRiskScore;
		return riskScore;
	}

	private String getOverallRiskClassification(int riskScore) {
		String overallRisk = null;

		if (IntStream.rangeClosed(0, 0).boxed().collect(Collectors.toList()).contains(riskScore))
			overallRisk = "No Risk";
		if (IntStream.rangeClosed(1, 6).boxed().collect(Collectors.toList()).contains(riskScore))
			overallRisk = "Low Risk";
		if (IntStream.rangeClosed(7, 13).boxed().collect(Collectors.toList()).contains(riskScore))
			overallRisk = "Medium Risk";
		if (IntStream.rangeClosed(14, 24).boxed().collect(Collectors.toList()).contains(riskScore))
			overallRisk = "High Risk";
		if (IntStream.rangeClosed(25, 34).boxed().collect(Collectors.toList()).contains(riskScore))
			overallRisk = "Critical Risk";

		return overallRisk;
	}

}
/*
 * Cell cell = row.getCell(6); System.out.println("Cell================="+cell);
 * if (cell == null) { // This cell is empty/blank/un-used, handle as needed
 * dataSheet.setRamlReviewDate(null); } else {
 * dataSheet.setRamlReviewDate(row.getCell(6).getDateCellValue()); // String
 * cellStr = fmt.formatCell(cell); // Do something with the value }
 */