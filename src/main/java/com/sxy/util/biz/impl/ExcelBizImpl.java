package com.sxy.util.biz.impl;

import com.sxy.util.bean.FormBean;
import com.sxy.util.biz.AbstractBiz;
import com.sxy.util.biz.IExcelBiz;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;

/**
 * @author weitao1
 * @time 2020/02/17
 */
@Service
public class ExcelBizImpl extends AbstractBiz implements IExcelBiz {

    @Override
    public InputStream avg(InputStream fis) {

        // 设置数字格式
        jxl.write.NumberFormat nf = new jxl.write.NumberFormat("#0");
        jxl.write.WritableCellFormat wcfN = new jxl.write.WritableCellFormat(nf);
        Map<String, FormBean> dataMap = new HashMap<String, FormBean>(600, 1);
        List<String> keys = new ArrayList<>();

        // 获得工作簿对象
        Workbook originalSheet = null;
        try {
            originalSheet = Workbook.getWorkbook(fis);
        } catch (IOException | BiffException e) {
            e.printStackTrace();
        }
        // 获得所有工作表
        Sheet[] sheets = Objects.requireNonNull(originalSheet).getSheets();
        // 遍历工作表
        if (sheets != null && sheets.length > 0) {
            Sheet sheet = sheets[0];
            // 获得行数
            int rows = sheet.getRows();
            // 读取数据
            for (int row = 1; row < rows; row++) {
                int col = 1;
                Cell cell = sheet.getCell(col, row);
                String id = cell.getContents().trim();
                Double automaticReview = Double.valueOf(sheet.getCell(2, row).getContents());
                Double returnVisit = Double.valueOf(sheet.getCell(3, row).getContents());
                Double reply = Double.valueOf(sheet.getCell(5, row).getContents());
                //dataMap 包含这个 id 将本次的值添加进去
                if (dataMap.containsKey(id)) {
                    dataMap.get(id).add(automaticReview, returnVisit, reply);
                } else {
                    //dataMap 不包含这个 id 创建一个form 传入参数,id 作为 key
                    dataMap.put(id, new FormBean(automaticReview, returnVisit, reply, 1d));
                    keys.add(id);
                }
            }

        }
        originalSheet.close();


        // 创建一个工作簿,承载目标工作表
        WritableWorkbook targetSheet;
        // 创建一个工作表,输出目标文件
        WritableSheet sheet;
        ByteArrayOutputStream os = null;
        try {
            os = new ByteArrayOutputStream();
            targetSheet = Workbook.createWorkbook(os);
            sheet = targetSheet.createSheet("sheet1", 0);
            //表头
            sheet.addCell(new Label(0, 0, "计划id"));
            sheet.addCell(new Label(1, 0, "自动评论人数"));
            sheet.addCell(new Label(2, 0, "回访人数"));
            sheet.addCell(new Label(3, 0, "回访率"));
            sheet.addCell(new Label(4, 0, "回复人数"));
            sheet.addCell(new Label(5, 0, "回复率"));

            int size = dataMap.size();
            for (int row = 1; row <= size; row++) {
                String keyId = keys.get(row - 1);
                FormBean bean = dataMap.get(keyId);
                Double cnt = bean.getCnt();
                sheet.addCell(new Label(0, row, keyId));
                sheet.addCell(new jxl.write.Number(1, row, bean.getAutomaticReview() / cnt, wcfN));
                sheet.addCell(new jxl.write.Number(2, row, bean.getReturnVisit() / cnt, wcfN));
                sheet.addCell(new Label(3, row, ""));
                sheet.addCell(new jxl.write.Number(4, row, bean.getReply() / cnt, wcfN));
                sheet.addCell(new Label(5, row, ""));
            }
            targetSheet.write();
            targetSheet.close();
        } catch (WriteException | IOException e) {
            e.printStackTrace();
        }
        byte[] byteArray = os.toByteArray();
        return new ByteArrayInputStream(byteArray);
    }
}

