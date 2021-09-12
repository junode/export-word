package com.demo.exportword.java2word;

import com.demo.exportword.java2word.interfaces.IDocument;
import com.demo.exportword.java2word.w2004.Document2004;
import com.demo.exportword.java2word.w2004.elements.BreakLine;
import com.demo.exportword.java2word.w2004.elements.Heading1;
import com.demo.exportword.java2word.w2004.elements.Paragraph;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RequestMapping;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.PrintWriter;

/**
 * demo 来源： https://www.jianshu.com/p/f773889bf627?utm_campaign=maleskine&utm_content=note&utm_medium=seo_notes&utm_source=recommendation
 * <p>
 * 目标：将 String 输出到 word文档
 * <p>
 * 访问：http://localhost:11111/javaToWord/export
 */
@CrossOrigin("*")
@Controller
@RequestMapping("/javaToWord")
public class JavaToWordController {

    public static void main(String[] args) {
        IDocument myDoc = new Document2004();
        myDoc.addEle(Heading1.with("昕").create());
        myDoc.addEle(BreakLine.times(2).create()); //two break lines
        myDoc.addEle(Paragraph.with("月亮代表我的心").create());
        //then get the XML representation of the MS Word document
        String myWord = myDoc.getContent();
        System.out.println(myWord);

        File fileObj = new File("E:\\github\\ruoyi\\word_file\\Java2word_allInOne.doc");
        //将生成的xml内容写入到文件中。
        PrintWriter writer = null;
        try {
            writer = new PrintWriter(fileObj);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        writer.println(myWord);
        writer.close();
    }

    @RequestMapping("/export")
    public ResponseEntity<Object> export() {
        IDocument myDoc = new Document2004();
        myDoc.addEle(Heading1.with("昕").create());
        myDoc.addEle(BreakLine.times(2).create()); //two break lines
        myDoc.addEle(Paragraph.with("月亮代表我的心").create());
        //then get the XML representation of the MS Word document
        String myWord = myDoc.getContent();

        byte[] a = myWord.getBytes();
        String filename = "my.docx";
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "fileName=\"" + filename + "\"")
                //文件名称
                .header(HttpHeaders.CONTENT_TYPE, "application/msword")
                //数据类型，这里也可用application/octet-stream 表示是二进制流
                .header(HttpHeaders.CONTENT_LENGTH, a.length + "").header("Connection", "close")
                //数据部分长度，即文件字节数组的长度
                .body(a);
                //写入文件内容(byte数组),直接写入字符串也可
    }

}
