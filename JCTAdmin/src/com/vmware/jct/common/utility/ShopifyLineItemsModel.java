package com.vmware.jct.common.utility;

import java.util.List;

public class ShopifyLineItemsModel {
	private String fulfillment_service;
	private String fulfillment_status;
	private String gift_card;
	private String grams;
	private String id;
	private String price;
	private String product_id;
	private String quantity;
	private String requires_shipping;
	private String sku;
	private String taxable;
	private String title;
	private String variant_id;
	private String variant_title;
	private String vendor;
	private String name;
	private String variant_inventory_management;
	private List<ShopifyPropertiesModel> properties;
	private String product_exists;
	private String fulfillable_quantity;
	private String total_discount;
	private List<ShopifyTakLinesModel> tax_lines;
	public String getFulfillment_service() {
		return fulfillment_service;
	}
	public void setFulfillment_service(String fulfillment_service) {
		this.fulfillment_service = fulfillment_service;
	}
	public String getFulfillment_status() {
		return fulfillment_status;
	}
	public void setFulfillment_status(String fulfillment_status) {
		this.fulfillment_status = fulfillment_status;
	}
	public String getGift_card() {
		return gift_card;
	}
	public void setGift_card(String gift_card) {
		this.gift_card = gift_card;
	}
	public String getGrams() {
		return grams;
	}
	public void setGrams(String grams) {
		this.grams = grams;
	}
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public String getPrice() {
		return price;
	}
	public void setPrice(String price) {
		this.price = price;
	}
	public String getProduct_id() {
		return product_id;
	}
	public void setProduct_id(String product_id) {
		this.product_id = product_id;
	}
	public String getQuantity() {
		return quantity;
	}
	public void setQuantity(String quantity) {
		this.quantity = quantity;
	}
	public String getRequires_shipping() {
		return requires_shipping;
	}
	public void setRequires_shipping(String requires_shipping) {
		this.requires_shipping = requires_shipping;
	}
	public String getSku() {
		return sku;
	}
	public void setSku(String sku) {
		this.sku = sku;
	}
	public String getTaxable() {
		return taxable;
	}
	public void setTaxable(String taxable) {
		this.taxable = taxable;
	}
	public String getTitle() {
		return title;
	}
	public void setTitle(String title) {
		this.title = title;
	}
	public String getVariant_id() {
		return variant_id;
	}
	public void setVariant_id(String variant_id) {
		this.variant_id = variant_id;
	}
	public String getVariant_title() {
		return variant_title;
	}
	public void setVariant_title(String variant_title) {
		this.variant_title = variant_title;
	}
	public String getVendor() {
		return vendor;
	}
	public void setVendor(String vendor) {
		this.vendor = vendor;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getVariant_inventory_management() {
		return variant_inventory_management;
	}
	public void setVariant_inventory_management(String variant_inventory_management) {
		this.variant_inventory_management = variant_inventory_management;
	}
	public String getProduct_exists() {
		return product_exists;
	}
	public void setProduct_exists(String product_exists) {
		this.product_exists = product_exists;
	}
	public String getFulfillable_quantity() {
		return fulfillable_quantity;
	}
	public void setFulfillable_quantity(String fulfillable_quantity) {
		this.fulfillable_quantity = fulfillable_quantity;
	}
	public String getTotal_discount() {
		return total_discount;
	}
	public void setTotal_discount(String total_discount) {
		this.total_discount = total_discount;
	}
	public List<ShopifyPropertiesModel> getProperties() {
		return properties;
	}
	public void setProperties(List<ShopifyPropertiesModel> properties) {
		this.properties = properties;
	}
	public List<ShopifyTakLinesModel> getTax_lines() {
		return tax_lines;
	}
	public void setTax_lines(List<ShopifyTakLinesModel> tax_lines) {
		this.tax_lines = tax_lines;
	}
}
