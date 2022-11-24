package com.tjltech;

import com.tjltech.utils.PoiUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class Main {

    private static final String excelName = "test.xlsx";
    private static final boolean isTitleExist = true;//是否有表头
    private static final int defaultColumnNum = 0;
    private static final String BRACKETS = "'";

    public static void main(String[] args) {
        System.out.println("############################ excel处理开始");
        InputStream in = null;
        Workbook wb = null;
        try {
            File file = new File(excelName);
            in = new FileInputStream(file);
            wb = PoiUtil.getWorkbook(PoiUtil.EXCEL_XLSX, in);
        } catch (IOException e1) {
            throw new RuntimeException(e1);
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
        }

        // 处理第一个sheet
        Sheet sheet = wb.getSheetAt(0);
        if (sheet == null) {
            throw new RuntimeException("未找到导入文件");
        }

        // 读取数据的总条数
        Integer count = sheet.getLastRowNum();
        StringBuffer sb = new StringBuffer();
        List<String> list = new ArrayList<String>();
        sb.append("{");
        for (int i = (isTitleExist?1:0); i <= count; i++) {
            Row row = sheet.getRow(i);
            // 不为空才处理数据
            if (row != null) {
                String s = String.valueOf(row.getCell(defaultColumnNum)).trim();
                if(s == null || s.equals("") || s.equals("null")){
                    continue;
                }
                list.add(s);
                sb.append(BRACKETS);
                sb.append(s);
                sb.append(BRACKETS);
                sb.append(",");
            }
        }
        String str = sb.toString();
        if(list.size()>0){
            str = str.substring(0, str.length()-1);
        }
        str = str + "}";
        System.out.println("共有 " + list.size() + " 条数据");
        System.out.println("生成字符串： " + str);
        System.out.println("############################ excel处理结束");
    }
}