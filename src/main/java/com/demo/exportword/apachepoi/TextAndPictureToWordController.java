

package com.demo.exportword.apachepoi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.UUID;
import java.util.stream.Collectors;
import java.util.stream.Stream;


/**
 * demo来源：https://www.baeldung.com/java-microsoft-word-with-apache-poi
 * <p>
 * 示例：将三个txt文档和一个png图片合并为 MS word文档
 * <p>
 * http://localhost:11111/testAndPic/export
 */
@CrossOrigin("*")
@Controller
@RequestMapping("/testAndPic")
public class TextAndPictureToWordController {
    public static String dirPrefix = "E:\\github\\ruoyi\\word_file\\";
    public static String logo = "static/content/logo-leaf.png";
    public static String paragraph1 = "static/content/poi-word-para1.txt";
    public static String paragraph2 = "static/content/poi-word-para2.txt";
    public static String paragraph3 = "static/content/poi-word-para3.txt";
    public static String output = "rest-with-spring.docx";

    @RequestMapping("/export")
    public void exportWord(HttpServletRequest request, HttpServletResponse response) throws Exception {
        System.out.println("get access .... ");
        XWPFDocument document = new XWPFDocument();
        // 1、格式化标题和子标题
        // 1.1、标题
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("Build Your REST API with Spring");
        titleRun.setColor("009933");
        titleRun.setBold(true);
        titleRun.setFontFamily("Courier");
        titleRun.setFontSize(20);

        // 1.2、子标题
        XWPFParagraph subTitle = document.createParagraph();
        subTitle.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun subTitleRun = subTitle.createRun();
        subTitleRun.setText("spring is not easy. Be careful");
        subTitleRun.setColor("00CC44");
        subTitleRun.setFontFamily("Courier");
        subTitleRun.setFontSize(16);
        subTitleRun.setTextPosition(20);
        subTitleRun.setUnderline(UnderlinePatterns.DOT_DOT_DASH);

        // 2、新增图片
        XWPFParagraph image = document.createParagraph();
        image.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun imageRun = image.createRun();
        imageRun.setTextPosition(20);

        Path imagePath = Paths.get(ClassLoader.getSystemResource(logo).toURI());
        imageRun.addPicture(Files.newInputStream(imagePath),
                XWPFDocument.PICTURE_TYPE_PNG, imagePath.getFileName().toString(),
                Units.toEMU(50), Units.toEMU(50));

        // 3、新增三段文本
        XWPFParagraph para1 = document.createParagraph();
        para1.setAlignment(ParagraphAlignment.BOTH);
        String string1 = convertTextFileToString(paragraph1);
        XWPFRun para1Run = para1.createRun();
        para1Run.setText(string1);

        XWPFParagraph para2 = document.createParagraph();
        para2.setAlignment(ParagraphAlignment.RIGHT);
        String string2 = convertTextFileToString(paragraph2);
        XWPFRun para2Run = para2.createRun();
        para2Run.setText(string2);
        para2Run.setItalic(true);

        XWPFParagraph para3 = document.createParagraph();
        para3.setAlignment(ParagraphAlignment.LEFT);
        String string3 = convertTextFileToString(paragraph3);
        XWPFRun para3Run = para3.createRun();
        para3Run.setText(string3);

        //设置文件名
        String fileName = UUID.randomUUID().toString();
        //处理文件名乱码
        String userAgent = request.getHeader("User-Agent");
        // 针对IE或者以IE为内核的浏览器：
        if (userAgent.contains("MSIE") || userAgent.contains("Trident")) {
            fileName = java.net.URLEncoder.encode(fileName, "UTF-8");
        } else {
            // 非IE浏览器
            fileName = new String(fileName.getBytes("UTF-8"), "ISO-8859-1");
        }
        // 4、导出
        request.setCharacterEncoding("utf-8");
        response.setContentType("application/msword");
        response.addHeader("Content-Disposition", "attachment;filename=" + fileName + ".doc");
        OutputStream ostream = response.getOutputStream();
        document.write(ostream);
        document.close();
        ostream.close();
    }

    public static String convertTextFileToString(String fileName) {
        try (Stream<String> stream
                     = Files.lines(Paths.get(ClassLoader.getSystemResource(fileName).toURI()))) {

            return stream.collect(Collectors.joining(" "));
        } catch (IOException | URISyntaxException e) {
            return null;
        }
    }

    @Test
    public void test1() throws IOException, InvalidFormatException, URISyntaxException {
        XWPFDocument document = new XWPFDocument();

        // 1、格式化标题和子标题
        // 1.1、标题
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("Build Your REST API with Spring");
        titleRun.setColor("009933");
        titleRun.setBold(true);
        titleRun.setFontFamily("Courier");
        titleRun.setFontSize(20);

        // 1.2、子标题
        XWPFParagraph subTitle = document.createParagraph();
        subTitle.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun subTitleRun = subTitle.createRun();
        subTitleRun.setText("spring is not easy. Be careful");
        subTitleRun.setColor("00CC44");
        subTitleRun.setFontFamily("Courier");
        subTitleRun.setFontSize(16);
        subTitleRun.setTextPosition(20);
        subTitleRun.setUnderline(UnderlinePatterns.DOT_DOT_DASH);

        // 2、新增图片
        XWPFParagraph image = document.createParagraph();
        image.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun imageRun = image.createRun();
        imageRun.setTextPosition(20);

        Path imagePath = Paths.get(ClassLoader.getSystemResource(logo).toURI());
        imageRun.addPicture(Files.newInputStream(imagePath),
                XWPFDocument.PICTURE_TYPE_PNG, imagePath.getFileName().toString(),
                Units.toEMU(50), Units.toEMU(50));

        // 3、新增三段文本
        XWPFParagraph para1 = document.createParagraph();
        para1.setAlignment(ParagraphAlignment.BOTH);
        String string1 = convertTextFileToString(paragraph1);
        XWPFRun para1Run = para1.createRun();
        para1Run.setText(string1);

        XWPFParagraph para2 = document.createParagraph();
        para2.setAlignment(ParagraphAlignment.RIGHT);
        String string2 = convertTextFileToString(paragraph2);
        XWPFRun para2Run = para2.createRun();
        para2Run.setText(string2);
        para2Run.setItalic(true);

        XWPFParagraph para3 = document.createParagraph();
        para3.setAlignment(ParagraphAlignment.LEFT);
        String string3 = convertTextFileToString(paragraph3);
        XWPFRun para3Run = para3.createRun();
        para3Run.setText(string3);

        // 4、导出
        FileOutputStream out = new FileOutputStream(dirPrefix + output);
        document.write(out);
        out.close();
        document.close();
    }
}