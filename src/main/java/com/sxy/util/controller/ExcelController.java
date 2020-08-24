package com.sxy.util.controller;

import com.sxy.util.biz.IExcelBiz;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author weitao1
 * @time 2020/02/18
 */
@RequestMapping("/excel")
@Controller
public class ExcelController {

    @Autowired
    private IExcelBiz excelBiz;

    @RequestMapping("/avg")
    public InputStream getExcelAvg(HttpServletResponse response, MultipartFile file) {
        InputStream is = null;
        try {
            is = excelBiz.avg(file.getInputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
        response.setHeader("Content-disposition", "attachment; filename=" + System.currentTimeMillis() + ".xls");
        response.setContentType("application/msexcel");
        return is;
    }

//    @RequestMapping("/cihui")
//    public InputStream getExcelLetter(HttpServletResponse response){
//        InputStream is = null;
//        is = excelBiz.randomLetter();
//        response.setHeader("Content-disposition", "attachment; filename=" + System.currentTimeMillis() + ".xls");
//        response.setContentType("application/msexcel");
//        return is;
//    }
    @RequestMapping("/test")
    public String test(){
        System.out.println("test");
        return "index.html";
    }

}
