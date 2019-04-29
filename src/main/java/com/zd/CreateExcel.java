package com.zd;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @Auther: Administrator
 * @Date: 2019/4/10 15:47
 * @Description:
 */
public class CreateExcel {
    public static void main(String args[]) {
        CreateExcel createExcel = new CreateExcel();
        String path = "C:\\Users\\Administrator\\Desktop\\test.xlsx";//路径可随意替换
        try {
            //随意创建一个Excel
            createExcel.createExcel(path);
            //读取上一行创建的Excel
            createExcel.getExcel(path);
            System.out.println("----------我是分割线----------");
            //创建Excel的表头
            createExcel.createExcelTop(path);
            //读取上一行创建的Excel
            //createExcel.getExcel(path);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    //创建Excel文件
    public void createExcel(String path) throws Exception {
        //创建Excel文件对象
        HSSFWorkbook wb = new HSSFWorkbook();
        //用文件对象创建sheet对象  
        HSSFSheet sheet = wb.createSheet("第一个sheet页");
        //用sheet对象创建行对象  
        HSSFRow row = sheet.createRow(0);
        //创建单元格样式     
        CellStyle cellStyle = wb.createCellStyle();
        //用行对象创建单元格对象Cell 
        Cell cell = row.createCell(0);
        //用cell对象读写。设置Excel工作表的值
        cell.setCellValue(1);
        //输出Excel文件
        FileOutputStream output = new FileOutputStream(path);
        wb.write(output);
        output.flush();
    }

    //读取Excel文件的值
    public void getExcel(String path) throws Exception {

        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(path));
        //得到Excel工作簿对象    
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        //得到sheet页个数（从1开始数,但是取值的时候要从index=0开始）
        int scount = wb.getNumberOfSheets();
        System.out.println("sheet页的个数为：" + (scount));
        for (int a = 0; a < scount; a++) {
            String sheetName = wb.getSheetName(a);
            System.out.println("第" + (a + 1) + "个sheet页的名字为" + sheetName + ",内容如下：");
            //得到Excel工作表对象(0代表第一个sheet页)    
            HSSFSheet sheet = wb.getSheetAt(a);
            HSSFSheet sheet1 = wb.getSheet("第一个sheet页");
            //预定义单元格的值
            String c = "";
            //得到工作表的有效行数(行数从0开始数，取值时从index=0开始)
            int rcount = sheet.getLastRowNum();
            System.out.println("第" + (a + 1) + "个sheet页有" + rcount + "行");
            for (int i = 0; i <= rcount; i++) {
                //得到Excel工作表的行    
                HSSFRow row = sheet.getRow(i);
                if (null != row) {
                    //获取一行（row）的有效单元格（cell）个数(列数从1开始数，取值的时候从index=0开始取)
                    int ccount = row.getLastCellNum();
                    System.out.println("第" + (i + 1) + "行有" + ccount + "个单元格");
                    for (int j = 0; j < ccount; j++) {
                        //得到Excel工作表指定行的单元格    
                        HSSFCell cell = row.getCell(j);
                        if (null != cell) {
                            //得到单元格类型
                            int cellType = cell.getCellType();
                            switch (cellType) {
                                //字符串类型  
                                case HSSFCell.CELL_TYPE_STRING:
                                    c = cell.getStringCellValue();
                                    if (c.trim().equals("") || c.trim().length() <= 0)
                                        c = " ";
                                    break;
                                case HSSFCell.CELL_TYPE_NUMERIC:
                                    c = String.valueOf(cell.getNumericCellValue());
                                default:
                                    break;
                            }
                            //String c = cell.getStringCellValue();
                            System.out.print("第" + (i + 1) + "行" + (j + 1) + "列的值为：" + c + "    ");
                        } else {
                            System.out.print("第" + (i + 1) + "行" + (j + 1) + "列的值为：空" + "    ");
                        }
                    }
                    System.out.println();
                }else {
                    System.out.println("第" + (i + 1) + "行的值为：空");
                }
            }
        }
    }
    //创建excel的表头,设置字体及字体颜色，设置单元格填充色
    public void createExcelTop (String path) throws Exception{
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        //创建Excel对象
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        //创建sheet对象
        HSSFSheet sheet = hssfWorkbook.createSheet();
        sheet.setDefaultRowHeight((short)200);
        //Excel样式
        HSSFCellStyle style = hssfWorkbook.createCellStyle();
        //Excel字体设置
        HSSFFont hssfFont = hssfWorkbook.createFont();
        hssfFont.setBoldweight(HSSFFont.COLOR_NORMAL);
        hssfFont.setFontHeightInPoints((short)13);
        hssfFont.setFontName("SimSun");
        style.setFont(hssfFont);
        //水平居中
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //垂直居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        HSSFFont Font = hssfWorkbook.createFont();
        Font.setFontHeightInPoints((short)9);
        Font.setFontName("SimSun");

        HSSFCellStyle fontRightStyle = hssfWorkbook.createCellStyle();
        fontRightStyle.setFont(Font);
        fontRightStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);

        HSSFCellStyle fontLeftStyle = hssfWorkbook.createCellStyle();
        fontLeftStyle.setFont(Font);
        fontLeftStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);


        HSSFCellStyle fontCenterStyle = hssfWorkbook.createCellStyle();
        fontCenterStyle.setFont(Font);
        fontCenterStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        fontCenterStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        HSSFCellStyle borderStyle = hssfWorkbook.createCellStyle();
        borderStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        borderStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        borderStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        borderStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);

        /**
         * 第一行
         * */
        //设置标题，以及设置单元格的样式，有些样式只对单元格有效
        //合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        Row row1 = sheet.createRow(0);
        row1.setHeight((short)1000);
        row1.createCell(0).setCellValue(" ");
        Cell cell1 = row1.createCell(1);
        cell1.setCellValue("同济医院 价值申请变动单");
        cell1.setCellStyle(style);
        CellRangeAddress address = new CellRangeAddress(0,0,1,8);
        sheet.addMergedRegion(address);

        /**
         * 第二行
         * */
        //样式

        CellRangeAddress address2 = new CellRangeAddress(1,1,2,4);
        sheet.addMergedRegion(address2);
        HSSFRow row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue(" ");
        HSSFCell row2Cell1 = row2.createCell(1);
        row2Cell1.setCellValue("科室:");
        row2Cell1.setCellStyle(fontRightStyle);
        HSSFCell row2Cell2 = row2.createCell(2);
        row2Cell2.setCellValue("总务科(主院区) 1086");
        row2Cell2.setCellStyle(fontLeftStyle);

        /**
         * 第三行
         * */
        HSSFRow row3 = sheet.createRow(2);
        row3.createCell(0).setCellValue(" ");
        row3.createCell(1).setCellValue("资产类别：");
        row3.createCell(2).setCellValue("土地、房屋及建筑物");
        row3.createCell(3);
        row3.createCell(4).setCellValue("资产子类：");
        row3.createCell(5).setCellValue("房屋");
        row3.createCell(6).setCellValue("制表日期：");
        row3.createCell(7).setCellValue(sdf.format(new Date()));
        row3.setRowStyle(fontCenterStyle);
        for(int rowIndex = 0;rowIndex<8;rowIndex++){
            if(rowIndex==2 || rowIndex==5 || rowIndex==7){
                row3.getCell(rowIndex).setCellStyle(fontLeftStyle);
            }else{
                row3.getCell(rowIndex).setCellStyle(fontRightStyle);
            }
            row3.getCell(rowIndex).getCellStyle().setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        }
      /*  row3.getCell(1).setCellStyle(fontRightStyle);
        row3.getCell(2).setCellStyle(fontLeftStyle);
        row3.getCell(4).setCellStyle(fontRightStyle);
        row3.getCell(5).setCellStyle(fontLeftStyle);
        row3.getCell(6).setCellStyle(fontRightStyle);
        row3.getCell(7).setCellStyle(fontLeftStyle);*/
        CellRangeAddress addressRow3 = new CellRangeAddress(2,2,2,3);
        sheet.addMergedRegion(addressRow3);
        /***
         * 第四行
         * */
        //创建row
        HSSFRow row = sheet.createRow(3);
        //设置表头
        HSSFCell rowCell1 = row.createCell(0);
        rowCell1.setCellValue("序号");
        row.createCell(1).setCellValue("资产编号");
        row.createCell(2).setCellValue("名称");
        row.createCell(3).setCellValue("变动前金额");
        row.createCell(4).setCellValue("变动金额");
        row.createCell(5).setCellValue("变动后金额");
        row.createCell(6).setCellValue("变动原因");
        row.createCell(7).setCellValue("变动日期");
        row.createCell(8).setCellValue("变动单号");
        row.createCell(9).setCellValue("备注");

        HSSFCellStyle songFontStyle = hssfWorkbook.createCellStyle();
        HSSFFont songFont = hssfWorkbook.createFont();
        songFont.setFontHeightInPoints((short)11);
        songFont.setFontName("宋体");
        songFontStyle.setFont(songFont);
        songFontStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        songFontStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        songFontStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        songFontStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        songFontStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        songFontStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        rowCell1.setCellStyle(songFontStyle);

        //Excel设置单元格填充色
        HSSFCellStyle style2 = hssfWorkbook.createCellStyle();
        style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);//设置前景填充样式
        style2.setFillForegroundColor(HSSFColor.RED.index);//前景填充色
        //设置数据
        for (int i = 4;i<6;i++){
            HSSFRow dateRow = sheet.createRow(i);
            dateRow.createCell(0).setCellValue(i-3);
            dateRow.createCell(1).setCellValue("FW20140191");
            dateRow.createCell(2).setCellValue("蔡甸专家社区C208-104");
            dateRow.createCell(3).setCellValue(300);
            HSSFCell changeCellSub = dateRow.createCell(4);
            String sub = "F"+(i+1)+"-D"+(i+1);
            changeCellSub.setCellFormula(sub);
            dateRow.createCell(5).setCellValue(200);
            dateRow.createCell(6).setCellValue("装修增值");
            dateRow.createCell(7).setCellValue(sdf.format(new Date()));
            dateRow.createCell(8).setCellValue("122398");
            dateRow.createCell(9).setCellValue("1装修增值");
            dateRow.getCell(0).setCellStyle(songFontStyle);
        }
        //合计行的开始行数
        int hjIndex = 6;
        /***
         * 合计行
         * */
        //创建row
        HSSFRow rowHJ = sheet.createRow(hjIndex);
        HSSFCell HJCell = rowHJ.createCell(0);
        HJCell.setCellValue("合计");
        HJCell.setCellStyle(songFontStyle);
        rowHJ.createCell(1);
        rowHJ.createCell(2);
        HSSFCell HJCellBefore = rowHJ.createCell(3);
        HJCellBefore.setCellFormula("SUM(D5:D6)");
        HSSFCell HJCellSub = rowHJ.createCell(4);
        HJCellSub.setCellFormula("F7-D7");
        CellRangeAddress addressHJ = new CellRangeAddress(hjIndex,hjIndex,0,1);
        sheet.addMergedRegion(addressHJ);

        HSSFCell HJCellAfter = rowHJ.createCell(5);
        HJCellAfter.setCellFormula("SUM(F5:F6)");
        sheet.setForceFormulaRecalculation(true);

        rowHJ.createCell(6);
        rowHJ.createCell(7);
        rowHJ.createCell(8);
        rowHJ.createCell(9);
        /***
         * 预算行
         * */
        HSSFRow rowYS = sheet.createRow(hjIndex+1);
        HSSFCell rowYSCell = rowYS.createCell(0);
        rowYSCell.setCellValue("预算信息");
        rowYSCell.setCellStyle(songFontStyle);
        CellRangeAddress addressYS = new CellRangeAddress(hjIndex+1,hjIndex+1,0,1);
        sheet.addMergedRegion(addressYS);
        CellRangeAddress addressYSDetail = new CellRangeAddress(hjIndex+1,hjIndex+1,2,7);
        sheet.addMergedRegion(addressYSDetail);
        createCell(rowYS,10,"预算信息明细项");

        /**
         * 操作员
         * */
        HSSFRow rowCZY = sheet.createRow(hjIndex+2);
        HSSFCell rowCZYCells = rowCZY.createCell(0);
        rowCZYCells.setCellValue("操作员");
        rowCZYCells.setCellStyle(songFontStyle);
        CellRangeAddress addressCZY= new CellRangeAddress(hjIndex+2,hjIndex+2,0,1);
        sheet.addMergedRegion(addressCZY);
        CellRangeAddress addressCZYName = new CellRangeAddress(hjIndex+2,hjIndex+2,2,7);
        sheet.addMergedRegion(addressCZYName);
        createCell(rowCZY,10,"操作员张三");


        /**
         * 签名
         * */
        HSSFRow rowSign = sheet.createRow(hjIndex+3);
        rowSign.createCell(1).setCellValue("处长：");
        rowSign.createCell(3).setCellValue("科长：");
        rowSign.createCell(5).setCellValue("保管：");
        rowSign.createCell(7).setCellValue("经手：");
        rowSign.setHeight((short)700);
        rowSign.getCell(1).setCellStyle(fontCenterStyle);
        for(int i=1;i<7;i+=2){
            rowSign.getCell(i+2).setCellStyle(fontCenterStyle);
        }


        for(int rowindex=3;rowindex<9;rowindex++){//行
            HSSFRow borderRow = sheet.getRow(rowindex);
            for(int column=0;column<10;column++){
               borderRow.getCell(column).setCellStyle(borderStyle);
               borderRow.getCell(column).getCellStyle().setFont(Font);
               borderRow.getCell(column).getCellStyle().setVerticalAlignment(CellStyle.VERTICAL_CENTER);
               borderRow.getCell(column).getCellStyle().setAlignment(CellStyle.ALIGN_CENTER);
            }
        }
        //设置列宽
        sheet.setColumnWidth(0,256*5);
        sheet.setColumnWidth(1,256*10);
        sheet.setColumnWidth(2,512*10);
        //输出Excel对象
        FileOutputStream output = new FileOutputStream(path);
        hssfWorkbook.write(output);
        output.flush();
    }
    /**
      * * 功能描述: *
      * @param:
      * @return:
      * @auther: Administrator
      * @date: 2019/4/11 14:16
     */
    public void createCell(HSSFRow row,int cellNum,String cellVal){
        for(int rowIndex=1;rowIndex<cellNum;rowIndex++){
            if(rowIndex==2){
                row.createCell(2).setCellValue(cellVal);
            }else{
                row.createCell(rowIndex);
            }
        }
    }
}
