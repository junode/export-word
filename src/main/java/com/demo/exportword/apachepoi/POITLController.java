package com.demo.exportword.apachepoi;

/*import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.util.PoitlIOUtils;*/

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.HashMap;

/**
 * 采用模板生成 WORD文档
 * 使用非常简单，但可惜这边demo添加maven依赖总是报错。
 */
@CrossOrigin("*")
@Controller
@RequestMapping("/poitl")
public class POITLController {

    @RequestMapping("/genByPOITL")
    public void genWordByPOITL(@RequestBody String content, HttpServletRequest request, HttpServletResponse response) throws IOException {
        /*System.out.println(content);
        XWPFTemplate template = XWPFTemplate.compile("template.docx").render(
                new HashMap<String, Object>() {{
                    put("title",  request.getParameter("content"));
                }});

        // 写入到网络流
        response.setContentType("application/octet-stream");
        response.setHeader("Content-disposition","attachment;filename=\""+"out_template.docx"+"\"");
        // HttpServletResponse response
        OutputStream out = response.getOutputStream();
        BufferedOutputStream bos = new BufferedOutputStream(out);
        template.write(bos);
        bos.flush();
        out.flush();
        PoitlIOUtils.closeQuietlyMulti(template, bos, out);*/
    }
}
