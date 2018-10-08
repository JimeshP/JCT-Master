package com.vmware.jct.common.utility;

import java.util.List;

public class ShopifyModel {
	private String email;
	private String id;
	private String name;
	private String total_price;
	List<ShopifyNodeAttributeModel> note_attributes;
	List<ShopifyLineItemsModel> line_items;
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getTotal_price() {
		return total_price;
	}
	public void setTotal_price(String total_price) {
		this.total_price = total_price;
	}
	public List<ShopifyNodeAttributeModel> getNote_attributes() {
		return note_attributes;
	}
	public void setNote_attributes(List<ShopifyNodeAttributeModel> note_attributes) {
		this.note_attributes = note_attributes;
	}
	public List<ShopifyLineItemsModel> getLine_items() {
		return line_items;
	}
	public void setLine_items(List<ShopifyLineItemsModel> line_items) {
		this.line_items = line_items;
	}
	
	
}
