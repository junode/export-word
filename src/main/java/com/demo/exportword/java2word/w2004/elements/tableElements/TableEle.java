package com.demo.exportword.java2word.w2004.elements.tableElements;


public enum TableEle{	
	TH("th"), TD("td"), TF("tf"), TABLE_DEF("tableDef"); //TR("tr"),

	private String value;
	
	TableEle(String value){
		this.value = value;
	}
	
	public String getValue(){
		return this.value;
	}
	
};