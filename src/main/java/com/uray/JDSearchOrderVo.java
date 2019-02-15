package com.uray;

import java.io.Serializable;

public class JDSearchOrderVo implements Serializable {

	private static final long serialVersionUID = 1L;
	
	private Long loginId;
	private Long jdOrderId;
	private Integer orderDate;
	private Integer orderStatus;
	private Integer payStatus;
	private Integer freight;
	private Integer payPrice = 0;
	private Integer companyPaymony = 0;
	private String expandInfo;
	private Long sku;
	private String name;
	
	private String address;
	private Integer provinceId;
	private String provinceName;
	private Integer orderNum;
	
	public JDSearchOrderVo() {
		super();
	}
	
	public Long getLoginId() {
		return loginId;
	}
	public void setLoginId(Long loginId) {
		this.loginId = loginId;
	}
	public Long getJdOrderId() {
		return jdOrderId;
	}
	public void setJdOrderId(Long jdOrderId) {
		this.jdOrderId = jdOrderId;
	}
	public Integer getOrderDate() {
		return orderDate;
	}
	public void setOrderDate(Integer orderDate) {
		this.orderDate = orderDate;
	}
	public Integer getOrderStatus() {
		return orderStatus;
	}
	public void setOrderStatus(Integer orderStatus) {
		this.orderStatus = orderStatus;
	}
	public Integer getPayStatus() {
		return payStatus;
	}
	public void setPayStatus(Integer payStatus) {
		this.payStatus = payStatus;
	}
	public Integer getFreight() {
		return freight;
	}
	public void setFreight(Integer freight) {
		this.freight = freight;
	}
	public Integer getPayPrice() {
		return payPrice;
	}
	public void setPayPrice(Integer payPrice) {
		this.payPrice = payPrice;
	}
	public Integer getCompanyPaymony() {
		return companyPaymony;
	}
	public void setCompanyPaymony(Integer companyPaymony) {
		this.companyPaymony = companyPaymony;
	}
	public String getExpandInfo() {
		return expandInfo;
	}
	public void setExpandInfo(String expandInfo) {
		this.expandInfo = expandInfo;
	}
	public Long getSku() {
		return sku;
	}
	public void setSku(Long sku) {
		this.sku = sku;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getAddress() {
		return address;
	}
	public void setAddress(String address) {
		this.address = address;
	}
	public String getProvinceName() {
		return provinceName;
	}
	public void setProvinceName(String provinceName) {
		this.provinceName = provinceName;
	}
	public Integer getOrderNum() {
		return orderNum;
	}
	public void setOrderNum(Integer orderNum) {
		this.orderNum = orderNum;
	}

	public Integer getProvinceId() {
		return provinceId;
	}

	public void setProvinceId(Integer provinceId) {
		this.provinceId = provinceId;
	}

}
