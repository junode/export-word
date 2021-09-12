package com.demo.exportword.java2word.w2004;


import com.demo.exportword.java2word.interfaces.IBody;
import com.demo.exportword.java2word.interfaces.IElement;
import com.demo.exportword.java2word.interfaces.IFooter;
import com.demo.exportword.java2word.interfaces.IHeader;

public class Body2004 implements IBody {

	StringBuilder txt = new StringBuilder("");
	IHeader header = new Header2004();
	IFooter footer = new Footer2004();

	@Override
	public void addEle(IElement e) {
		this.txt.append("\n" + e.getContent());
	}

	@Override
	public void addEle(String str) {
		this.txt.append("\n" + str);
	}

	@Override
	public String getContent() {
		StringBuilder res = new StringBuilder();
		res.append("\n<w:body>");

		res.append(txt.toString());
		
		String header = this.getHeader().getContent();
		String footer = this.getFooter().getContent();
		if (!"".equals(header) || !"".equals(footer)){
			String header_footer_top = "<w:sectPr wsp:rsidR=\"00DB1FE5\" wsp:rsidSect=\"00471A86\">";
			String header_footer_botton = "</w:sectPr>";
			
			res.append("\n" + header_footer_top);
			res.append(header);//header has to be inside the w:body
			res.append(footer);//header has to be inside the w:body
			if (this.getHeader().getHideHeaderAndFooterFirstPage()){
				res.append(this.getHeader().getHideHeaderAndFooterFirstPageXml());
			}
			
			res.append("\n" + header_footer_botton);
		}
		
		res.append("\n</w:body>");
		return res.toString();
	}
	

	//### Getters and setters ###
	@Override
	public IHeader getHeader() {
		return header;
	}

	@Override
	public IFooter getFooter() {
		return footer;
	}

	
}
