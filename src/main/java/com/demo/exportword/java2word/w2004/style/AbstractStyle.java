package com.demo.exportword.java2word.w2004.style;

import com.demo.exportword.java2word.interfaces.IElement;
import com.demo.exportword.java2word.interfaces.ISuperStylin;

public abstract class AbstractStyle implements ISuperStylin {

	private IElement element;
	
	public void setElement(IElement element) {
		this.element = element;
	}
	
	public IElement create() {
		return this.element;
	}
	
}