package com.demo.exportword.java2word.interfaces;

public interface IHeader extends IHasElement{

	void setHideHeaderAndFooterFirstPage(boolean value);
	boolean getHideHeaderAndFooterFirstPage();
	
	String getHideHeaderAndFooterFirstPageXml();
	
}
