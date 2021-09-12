package com.demo.exportword.freemarker;

public class Table {
	private String id;
	private String name;
	private String img;
	
	public Table(String id, String name, String img) {
		super();
		this.id = id;
		this.name = name;
		this.img = img;
	}
	
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public String getImg() {
		return img;
	}
	public void setImg(String img) {
		this.img = img;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	
}
