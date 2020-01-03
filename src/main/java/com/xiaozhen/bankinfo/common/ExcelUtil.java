package com.xiaozhen.bankinfo.common;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;

/**
 * excel工具类
 * @author xz
 */
@SuppressWarnings("all")
public class ExcelUtil {

    public static void main(String[] args) {
        try(FileOutputStream out = new FileOutputStream("C:\\Users\\xz\\Desktop\\" + "xz.xls")) {
            data().write(out);
            out.flush();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    public static HSSFWorkbook data(){
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("数据统计");
        createTopTitle(workbook,sheet);
        createSecondTitle(workbook,sheet);
        createThirdTitle(workbook,sheet);
        testData(workbook,sheet,3);
        testData(workbook,sheet,4);
        testData(workbook,sheet,5);
        testData(workbook,sheet,6);
        testData(workbook,sheet,7);
        testData(workbook,sheet,8);
        testData(workbook,sheet,9);
        testData(workbook,sheet,10);
        return workbook;
    }

    private static void createTopTitle(HSSFWorkbook workbook,HSSFSheet sheet){
        HSSFRow titleRow = sheet.createRow(0);
        HSSFCell top1 = titleRow.createCell(5);
        top1.setCellValue("二级指标");
        top1.setCellStyle(titleCellStyle(workbook));
        CellRangeAddress top1Range = new CellRangeAddress(0, 0, 5, 18);
        sheet.addMergedRegion(top1Range);
        HSSFCell top2 = titleRow.createCell(19);
        top2.setCellValue("一级指标");
        top2.setCellStyle(titleCellStyle(workbook));
        CellRangeAddress top2Range = new CellRangeAddress(0, 0, 19, 23);
        sheet.addMergedRegion(top2Range);
    }
    private static void createSecondTitle(HSSFWorkbook workbook,HSSFSheet sheet){
        HSSFRow titleRow = sheet.createRow(1);
        HSSFCell t0 = titleRow.createCell(0);
        t0.setCellValue("网点编号");
        t0.setCellStyle(contentCellStyle(workbook));
        HSSFCell t1 = titleRow.createCell(1);
        t1.setCellValue("网点名称");
        t1.setCellStyle(contentCellStyle(workbook));
        HSSFCell t2 = titleRow.createCell(2);
        t2.setCellValue("所属银行名称");
        t2.setCellStyle(contentCellStyle(workbook));
        HSSFCell t3 = titleRow.createCell(3);
        t3.setCellValue("网点所处区域");
        t3.setCellStyle(contentCellStyle(workbook));
        HSSFCell t4 = titleRow.createCell(4);
        t4.setCellValue("调查日期");
        t4.setCellStyle(contentCellStyle(workbook));

        HSSFCell t5 = titleRow.createCell(5);
        t5.setCellValue("A1.残损币兑换服务有效性");
        t5.setCellStyle(contentCellStyle(workbook));
        HSSFCell t6 = titleRow.createCell(6);
        t6.setCellValue("A2.残损币兑换服务规范");
        t6.setCellStyle(contentCellStyle(workbook));
        HSSFCell t7 = titleRow.createCell(7);
        t7.setCellValue("A3.残损币兑换话术规范");
        t7.setCellStyle(contentCellStyle(workbook));
        HSSFCell t8 = titleRow.createCell(8);
        t8.setCellValue("A4.残损币兑换服务态度");
        t8.setCellStyle(contentCellStyle(workbook));

        HSSFCell t9 = titleRow.createCell(9);
        t9.setCellValue("B1.反宣币服务有效性");
        t9.setCellStyle(contentCellStyle(workbook));
        HSSFCell t10 = titleRow.createCell(10);
        t10.setCellValue("B2.反宣币兑换服务态度");
        t10.setCellStyle(contentCellStyle(workbook));

        HSSFCell t11 = titleRow.createCell(11);
        t11.setCellValue("C1.人名币真伪鉴别服务有效性");
        t11.setCellStyle(contentCellStyle(workbook));
        HSSFCell t12 = titleRow.createCell(12);
        t12.setCellValue("C2.人名币真伪鉴别服务态度");
        t12.setCellStyle(contentCellStyle(workbook));

        HSSFCell t13 = titleRow.createCell(13);
        t13.setCellValue("D2.现场小面额人名币兑换服务有效性");
        t13.setCellStyle(contentCellStyle(workbook));
        HSSFCell t14 = titleRow.createCell(14);
        t14.setCellValue("D3.预约小面额人名币兑换服务有效性");
        t14.setCellStyle(contentCellStyle(workbook));
        HSSFCell t15 = titleRow.createCell(15);
        t15.setCellValue("D4.预约小面额人名币兑换服务兑现情况");
        t15.setCellStyle(contentCellStyle(workbook));
        HSSFCell t16 = titleRow.createCell(16);
        t16.setCellValue("D5.零钱兑换服务有效性");
        t16.setCellStyle(contentCellStyle(workbook));
        HSSFCell t17 = titleRow.createCell(17);
        t17.setCellValue("D6.券别调剂话术规范性");
        t17.setCellStyle(contentCellStyle(workbook));

        HSSFCell t18 = titleRow.createCell(18);
        t18.setCellValue("E1.在行式ATM机人名币付出质量");
        t18.setCellStyle(contentCellStyle(workbook));

        HSSFCell t19 = titleRow.createCell(19);
        t19.setCellValue("残损币兑换服务");
        t19.setCellStyle(contentCellStyle(workbook));
        HSSFCell t20 = titleRow.createCell(20);
        t20.setCellValue("反宣币兑换服务");
        t20.setCellStyle(contentCellStyle(workbook));
        HSSFCell t21 = titleRow.createCell(21);
        t21.setCellValue("人民币真伪鉴别服务");
        t21.setCellStyle(contentCellStyle(workbook));
        HSSFCell t22 = titleRow.createCell(22);
        t22.setCellValue("券别调剂服务");
        t22.setCellStyle(contentCellStyle(workbook));
        HSSFCell t23 = titleRow.createCell(23);
        t23.setCellValue("在行式ATM机付出人民币质量调查");
        t23.setCellStyle(contentCellStyle(workbook));
        HSSFCell t24 = titleRow.createCell(24);
        t24.setCellValue("总分");
        t24.setCellStyle(contentCellStyle(workbook));
        HSSFCell t25 = titleRow.createCell(25);
        t25.setCellValue("调查人");
        t25.setCellStyle(contentCellStyle(workbook));

    }
    private static void createThirdTitle(HSSFWorkbook workbook,HSSFSheet sheet){
        HSSFRow titleRow = sheet.createRow(2);

        HSSFCell t5 = titleRow.createCell(5);
        t5.setCellValue("15");
        t5.setCellStyle(contentCellStyle(workbook));
        HSSFCell t6 = titleRow.createCell(6);
        t6.setCellValue("15");
        t6.setCellStyle(contentCellStyle(workbook));
        HSSFCell t7 = titleRow.createCell(7);
        t7.setCellValue("观察项");
        t7.setCellStyle(contentCellStyle(workbook));
        HSSFCell t8 = titleRow.createCell(8);
        t8.setCellValue("观察项");
        t8.setCellStyle(contentCellStyle(workbook));

        HSSFCell t9 = titleRow.createCell(9);
        t9.setCellValue("15");
        t9.setCellStyle(contentCellStyle(workbook));
        HSSFCell t10 = titleRow.createCell(10);
        t10.setCellValue("观察项");
        t10.setCellStyle(contentCellStyle(workbook));

        HSSFCell t11 = titleRow.createCell(11);
        t11.setCellValue("15");
        t11.setCellStyle(contentCellStyle(workbook));
        HSSFCell t12 = titleRow.createCell(12);
        t12.setCellValue("观察项");
        t12.setCellStyle(contentCellStyle(workbook));

        HSSFCell t13 = titleRow.createCell(13);
        t13.setCellValue("20");
        t13.setCellStyle(contentCellStyle(workbook));
        HSSFCell t14 = titleRow.createCell(14);
        t14.setCellValue("10");
        t14.setCellStyle(contentCellStyle(workbook));
        HSSFCell t15 = titleRow.createCell(15);
        t15.setCellValue("10");
        t15.setCellStyle(contentCellStyle(workbook));
        HSSFCell t16 = titleRow.createCell(16);
        t16.setCellValue("20");
        t16.setCellStyle(contentCellStyle(workbook));
        HSSFCell t17 = titleRow.createCell(17);
        t17.setCellValue("观察项");
        t17.setCellStyle(contentCellStyle(workbook));

        HSSFCell t18 = titleRow.createCell(18);
        t18.setCellValue("20");
        t18.setCellStyle(contentCellStyle(workbook));

        HSSFCell t19 = titleRow.createCell(19);
        t19.setCellValue("30");
        t19.setCellStyle(contentCellStyle(workbook));
        HSSFCell t20 = titleRow.createCell(20);
        t20.setCellValue("15");
        t20.setCellStyle(contentCellStyle(workbook));
        HSSFCell t21 = titleRow.createCell(21);
        t21.setCellValue("15");
        t21.setCellStyle(contentCellStyle(workbook));
        HSSFCell t22 = titleRow.createCell(22);
        t22.setCellValue("20");
        t22.setCellStyle(contentCellStyle(workbook));
        HSSFCell t23 = titleRow.createCell(23);
        t23.setCellValue("20");
        t23.setCellStyle(contentCellStyle(workbook));
        HSSFCell t24 = titleRow.createCell(24);
        t24.setCellValue("100");
        t24.setCellStyle(contentCellStyle(workbook));
    }

    private static void testData(HSSFWorkbook workbook,HSSFSheet sheet,int i){
        HSSFRow titleRow = sheet.createRow(i);
        HSSFCell t0 = titleRow.createCell(0);
        t0.setCellValue("11");
        t0.setCellStyle(contentCellStyle(workbook));
        HSSFCell t1 = titleRow.createCell(1);
        t1.setCellValue("青羊支行");
        t1.setCellStyle(contentCellStyle(workbook));
        HSSFCell t2 = titleRow.createCell(2);
        t2.setCellValue("建设银行");
        t2.setCellStyle(contentCellStyle(workbook));
        HSSFCell t3 = titleRow.createCell(3);
        t3.setCellValue("青羊区");
        t3.setCellStyle(contentCellStyle(workbook));
        HSSFCell t4 = titleRow.createCell(4);
        t4.setCellValue("2019-12-21");
        t4.setCellStyle(contentCellStyle(workbook));

        HSSFCell t5 = titleRow.createCell(5);
        t5.setCellValue("15");
        t5.setCellStyle(contentCellStyle(workbook));
        HSSFCell t6 = titleRow.createCell(6);
        t6.setCellValue("15");
        t6.setCellStyle(contentCellStyle(workbook));
        HSSFCell t7 = titleRow.createCell(7);
        t7.setCellValue("未使用规范话术");
        t7.setCellStyle(contentCellStyle(workbook));
        HSSFCell t8 = titleRow.createCell(8);
        t8.setCellValue("态度一般");
        t8.setCellStyle(contentCellStyle(workbook));

        HSSFCell t9 = titleRow.createCell(9);
        t9.setCellValue("15");
        t9.setCellStyle(contentCellStyle(workbook));
        HSSFCell t10 = titleRow.createCell(10);
        t10.setCellValue("态度一般");
        t10.setCellStyle(contentCellStyle(workbook));

        HSSFCell t11 = titleRow.createCell(11);
        t11.setCellValue("15");
        t11.setCellStyle(contentCellStyle(workbook));
        HSSFCell t12 = titleRow.createCell(12);
        t12.setCellValue("态度一般");
        t12.setCellStyle(contentCellStyle(workbook));

        HSSFCell t13 = titleRow.createCell(13);
        t13.setCellValue("20");
        t13.setCellStyle(contentCellStyle(workbook));
        HSSFCell t14 = titleRow.createCell(14);
        t14.setCellValue("10");
        t14.setCellStyle(contentCellStyle(workbook));
        HSSFCell t15 = titleRow.createCell(15);
        t15.setCellValue("10");
        t15.setCellStyle(contentCellStyle(workbook));
        HSSFCell t16 = titleRow.createCell(16);
        t16.setCellValue("20");
        t16.setCellStyle(contentCellStyle(workbook));
        HSSFCell t17 = titleRow.createCell(17);
        t17.setCellValue("使用规范话术");
        t17.setCellStyle(contentCellStyle(workbook));

        HSSFCell t18 = titleRow.createCell(18);
        t18.setCellValue("20");
        t18.setCellStyle(contentCellStyle(workbook));

        HSSFCell t19 = titleRow.createCell(19);
        t19.setCellValue("30");
        t19.setCellStyle(contentCellStyle(workbook));
        HSSFCell t20 = titleRow.createCell(20);
        t20.setCellValue("15");
        t20.setCellStyle(contentCellStyle(workbook));
        HSSFCell t21 = titleRow.createCell(21);
        t21.setCellValue("15");
        t21.setCellStyle(contentCellStyle(workbook));
        HSSFCell t22 = titleRow.createCell(22);
        t22.setCellValue("15");
        t22.setCellStyle(contentCellStyle(workbook));
        HSSFCell t23 = titleRow.createCell(23);
        t23.setCellValue("15");
        t23.setCellStyle(contentCellStyle(workbook));
        HSSFCell t24 = titleRow.createCell(24);
        t24.setCellValue("90");
        t24.setCellStyle(contentCellStyle(workbook));
        HSSFCell t25 = titleRow.createCell(25);
        t25.setCellValue("肖潇洒");
        t25.setCellStyle(contentCellStyle(workbook));

    }

    private static HSSFCellStyle contentCellStyle(HSSFWorkbook workbook) {
        HSSFCellStyle contentCellStyle = workbook.createCellStyle();
        contentCellStyle.setAlignment(HorizontalAlignment.CENTER);
        contentCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return contentCellStyle;
    }
    private static HSSFCellStyle titleCellStyle(HSSFWorkbook workbook){
        HSSFCellStyle titleCellStyle = workbook.createCellStyle();
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        HSSFFont titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleCellStyle.setFont(titleFont);
        return titleCellStyle;
    }

    public static void excelResponse(Workbook wb, HttpServletResponse response, String fileName) throws IOException {
        response.setHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes(), StandardCharsets.ISO_8859_1)+".xlsx");
        if(wb instanceof HSSFWorkbook){
            response.setHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes(), StandardCharsets.ISO_8859_1)+".xls");
        }
        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Cache-Control", "no-cache");
        response.setDateHeader("Expires", 0);
        OutputStream output = response.getOutputStream();
        BufferedOutputStream bufferedOutPut = new BufferedOutputStream(output);
        wb.write(bufferedOutPut);
        bufferedOutPut.flush();
        bufferedOutPut.close();
    }
}
