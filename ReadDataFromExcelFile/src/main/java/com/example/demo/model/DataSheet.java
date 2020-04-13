package com.example.demo.model;

import java.util.Date;

public class DataSheet {

	private int srNo;
	private String apiVersion;
	private String apiName;
	private String apiType;
	private String apiRiskClassificatin;
	private String ramlReviewStatus;
	private Date ramlReviewDate;
	private String veracodeStatus;
	private Date veracodeDate;
	private String penTestStatus;
	private Date penTestDate;
	private String veracodeSlaBreach;
	private String penTestSlaBreach;
	private String ramlReviewPending;
	private int riskScore;

	public DataSheet() {
		// TODO Auto-generated constructor stub
	}

	public DataSheet(int srNo, String apiVersion, String apiName, String apiType, String apiRiskClassificatin,
			String ramlReviewStatus, Date ramlReviewDate, String veracodeStatus, Date veracodeDate,
			String penTestStatus, Date penTestDate, String veracodeSlaBreach, String penTestSlaBreach,
			String ramlReviewPending, int riskScore) {
		super();
		this.srNo = srNo;
		this.apiVersion = apiVersion;
		this.apiName = apiName;
		this.apiType = apiType;
		this.apiRiskClassificatin = apiRiskClassificatin;
		this.ramlReviewStatus = ramlReviewStatus;
		this.ramlReviewDate = ramlReviewDate;
		this.veracodeStatus = veracodeStatus;
		this.veracodeDate = veracodeDate;
		this.penTestStatus = penTestStatus;
		this.penTestDate = penTestDate;
		this.veracodeSlaBreach = veracodeSlaBreach;
		this.penTestSlaBreach = penTestSlaBreach;
		this.ramlReviewPending = ramlReviewPending;
		this.riskScore = riskScore;
	}

	@Override
	public String toString() {
		return "DataSheet [srNo=" + srNo + ", apiVersion=" + apiVersion + ", apiName=" + apiName + ", apiType="
				+ apiType + ", apiRiskClassificatin=" + apiRiskClassificatin + ", ramlReviewStatus=" + ramlReviewStatus
				+ ", ramlReviewDate=" + ramlReviewDate + ", veracodeStatus=" + veracodeStatus + ", veracodeDate="
				+ veracodeDate + ", penTestStatus=" + penTestStatus + ", penTestDate=" + penTestDate
				+ ", veracodeSlaBreach=" + veracodeSlaBreach + ", penTestSlaBreach=" + penTestSlaBreach
				+ ", ramlReviewPending=" + ramlReviewPending + ", riskScore=" + riskScore + "]";
	}

	public int getSrNo() {
		return srNo;
	}

	public void setSrNo(int srNo) {
		this.srNo = srNo;
	}

	public String getApiVersion() {
		return apiVersion;
	}

	public void setApiVersion(String apiVersion) {
		this.apiVersion = apiVersion;
	}

	public String getApiName() {
		return apiName;
	}

	public void setApiName(String apiName) {
		this.apiName = apiName;
	}

	public String getApiType() {
		return apiType;
	}

	public void setApiType(String apiType) {
		this.apiType = apiType;
	}

	public String getApiRiskClassificatin() {
		return apiRiskClassificatin;
	}

	public void setApiRiskClassificatin(String apiRiskClassificatin) {
		this.apiRiskClassificatin = apiRiskClassificatin;
	}

	public String getRamlReviewStatus() {
		return ramlReviewStatus;
	}

	public void setRamlReviewStatus(String ramlReviewStatus) {
		this.ramlReviewStatus = ramlReviewStatus;
	}

	public Date getRamlReviewDate() {
		return ramlReviewDate;
	}

	public void setRamlReviewDate(Date ramlReviewDate) {
		this.ramlReviewDate = ramlReviewDate;
	}

	public String getVeracodeStatus() {
		return veracodeStatus;
	}

	public void setVeracodeStatus(String veracodeStatus) {
		this.veracodeStatus = veracodeStatus;
	}

	public Date getVeracodeDate() {
		return veracodeDate;
	}

	public void setVeracodeDate(Date veracodeDate) {
		this.veracodeDate = veracodeDate;
	}

	public String getPenTestStatus() {
		return penTestStatus;
	}

	public void setPenTestStatus(String penTestStatus) {
		this.penTestStatus = penTestStatus;
	}

	public Date getPenTestDate() {
		return penTestDate;
	}

	public void setPenTestDate(Date penTestDate) {
		this.penTestDate = penTestDate;
	}

	public String getVeracodeSlaBreach() {
		return veracodeSlaBreach;
	}

	public void setVeracodeSlaBreach(String veracodeSlaBreach) {
		this.veracodeSlaBreach = veracodeSlaBreach;
	}

	public String getPenTestSlaBreach() {
		return penTestSlaBreach;
	}

	public void setPenTestSlaBreach(String penTestSlaBreach) {
		this.penTestSlaBreach = penTestSlaBreach;
	}

	public String getRamlReviewPending() {
		return ramlReviewPending;
	}

	public void setRamlReviewPending(String ramlReviewPending) {
		this.ramlReviewPending = ramlReviewPending;
	}

	public int getRiskScore() {
		return riskScore;
	}

	public void setRiskScore(int riskScore) {
		this.riskScore = riskScore;
	}

	
	
	
}

