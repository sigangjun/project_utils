package cn.sigangjun.frame.poi.test;

import cn.sigangjun.frame.poi.ExcelTitleAnnotation;

public class DeviceDto {

	@ExcelTitleAnnotation("智能卡号")
	private String deviceId;
	@ExcelTitleAnnotation("序号")
	private Integer seq;
	@ExcelTitleAnnotation("业务号")
	private Integer serviceCode;
	@ExcelTitleAnnotation("1级机构")
	private String field1;
	@ExcelTitleAnnotation("2级机构")
	private String field2;
	@ExcelTitleAnnotation("3级机构")
	private String field3;
	@ExcelTitleAnnotation("4级机构")
	private String field4;
	@ExcelTitleAnnotation("5级机构")
	private String field5;
	@ExcelTitleAnnotation("6级机构")
	private String field6;
	@ExcelTitleAnnotation("7级机构")
	private String field7;
	@ExcelTitleAnnotation("8级机构")
	private String field8;

	public Integer getSeq() {
		return seq;
	}

	public void setSeq(Integer seq) {
		this.seq = seq;
	}

	public String getDeviceId() {
		return deviceId;
	}

	public void setDeviceId(String deviceId) {
		this.deviceId = deviceId;
	}

	public Integer getServiceCode() {
		return serviceCode;
	}

	public void setServiceCode(Integer serviceCode) {
		this.serviceCode = serviceCode;
	}

	public String getField1() {
		return field1;
	}

	public void setField1(String field1) {
		this.field1 = field1;
	}

	public String getField2() {
		return field2;
	}

	public void setField2(String field2) {
		this.field2 = field2;
	}

	public String getField3() {
		return field3;
	}

	public void setField3(String field3) {
		this.field3 = field3;
	}

	public String getField4() {
		return field4;
	}

	public void setField4(String field4) {
		this.field4 = field4;
	}

	public String getField5() {
		return field5;
	}

	public void setField5(String field5) {
		this.field5 = field5;
	}

	public String getField6() {
		return field6;
	}

	public void setField6(String field6) {
		this.field6 = field6;
	}

	public String getField7() {
		return field7;
	}

	public void setField7(String field7) {
		this.field7 = field7;
	}

	public String getField8() {
		return field8;
	}

	public void setField8(String field8) {
		this.field8 = field8;
	}

	@Override
	public String toString() {
		return "DeviceDto [deviceId=" + deviceId + ", seq=" + seq + ", serviceCode=" + serviceCode + ", field1=" + field1 + ", field2=" + field2 + ", field3=" + field3 + ", field4=" + field4 + ", field5=" + field5 + ", field6=" + field6 + ", field7=" + field7 + ", field8=" + field8 + "]";
	}

}
