package com.cmi.sonar;


import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class WriteExcelForXSSF {
    private static Log log = LogFactory.getLog(WriteExcelForXSSF.class);

    public void write(List<String> projectList, Map<String, Map<String, String>> dataMap, String startTime, String endTime) throws ParseException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("0");
        Row row = sheet.createRow(0);
        sheet.setColumnWidth(0, 40 * 256);
        CellStyle cellStyle = workbook.createCellStyle();
        // 设置这些样式
        cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        long m = sdf.parse(endTime).getTime() - sdf.parse(startTime).getTime();
        long day = m / (1000 * 60 * 60 * 24);
        //创建表头
        workbook.setSheetName(0, "AnalyzeReport");
        row.createCell(0).setCellStyle(cellStyle);
        row.createCell(0).setCellValue("ProjectName");
        for (int i = 0; i <= day; i++) {
            //转换时间
            DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
            Date dd = df.parse(startTime);
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(dd);
            calendar.add(Calendar.DAY_OF_MONTH, i);
            String date = df.format(calendar.getTime());
            //创建动态时间的列
            row.createCell(i + 1).setCellStyle(cellStyle);
            row.createCell(i + 1).setCellValue(date);
            sheet.setColumnWidth(i+1, 18 * 256);
        }

        //循环创建行
        for (int j = 0; j < projectList.size(); j++) {
            Row rowNum = sheet.createRow(j + 1);
            rowNum.createCell(0).setCellValue(projectList.get(j));
            for (int k=1;k<=day+1;k++){
                //转换时间
                DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
                Date dd = df.parse(startTime);
                Calendar calendar = Calendar.getInstance();
                calendar.setTime(dd);
                calendar.add(Calendar.DAY_OF_MONTH, k-1);
                String dateTime = df.format(calendar.getTime());
                Map<String,String>bugMap=dataMap.get(projectList.get(j));
                if (bugMap.get(dateTime)!=null&&bugMap.get(dateTime)!=""){
                    int num = Integer.valueOf(bugMap.get(dateTime));
                    rowNum.createCell(k).setCellValue(num);
                }
            }

        }

        try {
            Calendar cal=Calendar.getInstance();
            String date,daytime,month,year;
            year =String.valueOf(cal.get(Calendar.YEAR));
            month =String.valueOf(cal.get(Calendar.MONTH)+1);
            daytime =String.valueOf(cal.get(Calendar.DATE));
            date = year+"-"+month+"-"+daytime;
            System.out.println(date);
            File file = new File(".\\download\\AnalyzeReport_"+date+".xlsx");
            log.info("Report Path:"+ file.getPath());
            FileOutputStream fileoutputStream = new FileOutputStream(file);
            workbook.write(fileoutputStream);
            fileoutputStream.close();
        } catch (IOException e) {
            log.error("Export Report Error");
        }
    }
}