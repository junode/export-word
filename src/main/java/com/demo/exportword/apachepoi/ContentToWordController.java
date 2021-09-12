package com.demo.exportword.apachepoi;

import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.OutputStream;
import java.util.Objects;
import java.util.UUID;

/**
 * demo 来源： https://github.com/li-ze-lin/fck-export-word
 * <p>
 * 目标：将上传字符串转成word导出
 * <p>
 * 打开首页：http://localhost:1111/simple/index
 */
@CrossOrigin("*")
@Controller
@RequestMapping("/simple")
public class ContentToWordController {
    private volatile SimpleTextDto simpleText;

    @RequestMapping("/index")
    public String index() {
        return "simple";
    }

    @PostMapping("/save")
    public void save(@RequestBody SimpleTextDto simpleText) {
        this.simpleText = simpleText;
    }

    @RequestMapping("/export")
    public void exportWord(HttpServletRequest request, HttpServletResponse response) throws Exception {
        Objects.requireNonNull(simpleText);
        //word内容
        String content = "<html><meta charset=\"utf-8\" /><body>";
        content += simpleText.getText();
        content += "</body></html>";

        //设置编码的，不然导出中文就会乱码
        byte b[] = content.getBytes("utf-8");
        //将字节数组包装到流中
        ByteArrayInputStream bais = new ByteArrayInputStream(b);

        //* 生成word格式
        POIFSFileSystem poifs = new POIFSFileSystem();
        // 以下两行代码真不能少
        DirectoryEntry directory = poifs.getRoot();
        DocumentEntry documentEntry = directory.createDocument("WordDocument", bais);

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

        //输出文件
        request.setCharacterEncoding("utf-8");
        //导出word格式
        response.setContentType("application/msword");
        response.addHeader("Content-Disposition", "attachment;filename=" + fileName + ".doc");
        OutputStream ostream = response.getOutputStream();
        poifs.writeFilesystem(ostream);
        bais.close();
        ostream.close();
    }
}
