package com.demo.exportword.apachepoi;

/**
 * 富文本实体
 */
public class SimpleTextDto {

    private String text;

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public SimpleTextDto(String text) {
        this.text = text;
    }

    public SimpleTextDto() {
    }
}
