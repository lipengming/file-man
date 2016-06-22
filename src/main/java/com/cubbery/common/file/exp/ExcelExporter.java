/**
 * Copyright (c) 2016, www.cubbery.com. All rights reserved.
 */
package com.cubbery.common.file.exp;

import com.cubbery.common.file.Exporter;
import com.cubbery.common.file.FileInfo;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.*;
import com.cubbery.common.file.Header;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

import static com.cubbery.common.file.exp.CsvExporter.dateToStr;

/**
 * <b>类描述</b>：   <br>
 * <b>创建人</b>：   <a href="mailto:cubber.zh@gmail,com">百墨</a> <br>
 * <b>创建时间</b>： 2016/5/26 - 15:53  <br>
 *
 * @version 1.0.0   <br>
 */
public class ExcelExporter implements Exporter {

    @Override
    public void export(FileInfo data) throws IOException {
        ExcelHandler handler = new ExcelHandler();
        try {
            handler.export(data);
        } finally {
            handler.dispose();
        }
    }

    class ExcelHandler {
        //book
        private SXSSFWorkbook wb;
        // 标题行样式
        private CellStyle titleStyle;
        // 标题行字体
        private Font titleFont;
        // 日期行样式
        private CellStyle dateStyle;
        // 日期行字体
        private Font dateFont;
        // 表头行样式
        private CellStyle headStyle;
        // 表头行字体
        private Font headFont;
        // 内容行样式
        private CellStyle contentStyle;
        // 内容行字体
        private Font contentFont;

        public void export(FileInfo fileInfo) throws IOException {
            init();
            Sheet[] sheets = getSheets(1, new String[]{fileInfo.getSheetName() == null ? "sheet1" : fileInfo.getSheetName()});
            Sheet sheet = sheets[0];
            int headerIndex = fileInfo.getHeaderIndex();
            //表头
            createTableHeadRow(fileInfo, sheet, headerIndex);
            //表体
            createTableBodyRows(fileInfo, sheet, headerIndex + 1);
            //调整列宽
            adjustColumnSize(sheet,fileInfo.getHeaders().size());
            //写到输出流
            wb.write(fileInfo.getOutputStream());
        }

        /**
         * @Description: 初始化
         */
        private void init() {
            wb = new SXSSFWorkbook();
            titleFont = wb.createFont();
            titleStyle = wb.createCellStyle();
            dateStyle = wb.createCellStyle();
            dateFont = wb.createFont();
            headStyle = wb.createCellStyle();
            headFont = wb.createFont();
            contentStyle = wb.createCellStyle();
            contentFont = wb.createFont();

            initTitleCellStyle();
            initTitleFont();
            initDateCellStyle();
            initDateFont();
            initHeadCellStyle();
            initHeadFont();
            initContentCellStyle();
            initContentFont();
        }

        /**
         * @Description: 自动调整列宽
         */
        @SuppressWarnings("unused")
        private void adjustColumnSize(Sheet sheet, int columnSize) {
            for (int i = 0; i < columnSize + 1; i++) {
                sheet.autoSizeColumn(i, true);
            }
        }

        /**
         * @Description: 创建标题行(需合并单元格)
         */
        private void createTableTitleRow(FileInfo fileInfo, Sheet sheet, int rowNum) {
            CellRangeAddress titleRange = new CellRangeAddress(0, 0, 0, fileInfo.getHeaders().size());
            sheet.addMergedRegion(titleRange);
            Row titleRow = sheet.createRow(rowNum);
            titleRow.setHeight((short) 800);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellStyle(titleStyle);
            titleCell.setCellValue(fileInfo.getTitle());
        }

        /**
         * @Description: 创建日期行(需合并单元格)
         */
        private void createTableDateRow(FileInfo fileInfo, Sheet sheet, int rowNum) {
            CellRangeAddress dateRange = new CellRangeAddress(1, 1, 0, fileInfo.getHeaders().size());
            sheet.addMergedRegion(dateRange);
            Row dateRow = sheet.createRow(rowNum);
            dateRow.setHeight((short) 350);
            Cell dateCell = dateRow.createCell(0);
            dateCell.setCellStyle(dateStyle);
            dateCell.setCellValue(new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
        }

        /**
         * @Description: 创建表头行(需合并单元格)
         */
        private void createTableHeadRow(FileInfo fileInfo, Sheet sheet, int rowNum) {
            // 表头
            Row headRow = sheet.createRow(rowNum);
            headRow.setHeight((short) 350);
            // 列头名称
            for (Header header : fileInfo.getHeaders()) {
                Cell headCell = headRow.createCell(header.getColumnIndex());
                headCell.setCellStyle(headStyle);
                headCell.setCellValue(header.getColumnName());
            }
        }

        private void createTableBodyRows(FileInfo fileInfo, Sheet sheet, int rowNum) {
            int dataRow = rowNum;
            for (Map<String, String> data : fileInfo.getData()) {
                Row contentRow = sheet.createRow(dataRow);
                contentRow.setHeight((short) 300);
                Cell[] cells = getCells(contentRow, fileInfo.getHeaders().size());
                for (Header header : fileInfo.getHeaders()) {
                    String value = dateToStr(data.get(header.getFieldCode()));
                    cells[header.getColumnIndex()].setCellValue(value);
                }
                dataRow++;
            }
        }

        /**
         * @Description: 创建所有的Sheet
         */
        private Sheet[] getSheets(int num, String[] names) {
            Sheet[] sheets = new Sheet[num];
            for (int i = 0; i < num; i++) {
                sheets[i] = wb.createSheet(names[i]);
            }
            return sheets;
        }

        /**
         * @Description: 创建内容行的每一列
         */
        private Cell[] getCells(Row contentRow, int num) {
            Cell[] cells = new Cell[num];
            for (int i = 0, len = cells.length; i < len; i++) {
                cells[i] = contentRow.createCell(i);
                cells[i].setCellStyle(contentStyle);
            }
            return cells;
        }

        /**
         * @Description: 初始化标题行样式
         */
        private void initTitleCellStyle() {
            titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
            titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            titleStyle.setFont(titleFont);
            titleStyle.setFillBackgroundColor(IndexedColors.SKY_BLUE.getIndex());
        }

        /**
         * @Description: 初始化日期行样式
         */
        private void initDateCellStyle() {
            dateStyle.setAlignment(CellStyle.ALIGN_CENTER_SELECTION);
            dateStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            dateStyle.setFont(dateFont);
            dateStyle.setFillBackgroundColor(IndexedColors.SKY_BLUE.getIndex());
        }

        /**
         * @Description: 初始化表头行样式
         */
        private void initHeadCellStyle() {
            headStyle.setAlignment(CellStyle.ALIGN_CENTER);
            headStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            headStyle.setFont(headFont);
            headStyle.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
            headStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
            headStyle.setBorderBottom(CellStyle.BORDER_THIN);
            headStyle.setBorderLeft(CellStyle.BORDER_THIN);
            headStyle.setBorderRight(CellStyle.BORDER_THIN);
            headStyle.setTopBorderColor(IndexedColors.BLUE.getIndex());
            headStyle.setBottomBorderColor(IndexedColors.BLUE.getIndex());
            headStyle.setLeftBorderColor(IndexedColors.BLUE.getIndex());
            headStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());
        }

        /**
         * @Description: 初始化内容行样式
         */
        private void initContentCellStyle() {
            contentStyle.setAlignment(CellStyle.ALIGN_CENTER);
            contentStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            contentStyle.setFont(contentFont);
            contentStyle.setBorderTop(CellStyle.BORDER_THIN);
            contentStyle.setBorderBottom(CellStyle.BORDER_THIN);
            contentStyle.setBorderLeft(CellStyle.BORDER_THIN);
            contentStyle.setBorderRight(CellStyle.BORDER_THIN);
            contentStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            contentStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            contentStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            contentStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            contentStyle.setWrapText(true);    // 字段换行
        }

        /**
         * @Description: 初始化标题行字体
         */
        private void initTitleFont() {
            //titleFont.setFontName("华文楷体");
            titleFont.setFontHeightInPoints((short) 20);
            titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
            titleFont.setCharSet(Font.DEFAULT_CHARSET);
            titleFont.setColor(IndexedColors.BLACK.getIndex());
        }

        /**
         * @Description: 初始化日期行字体
         */
        private void initDateFont() {
            //dateFont.setFontName("隶书");
            dateFont.setFontHeightInPoints((short) 10);
            dateFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
            dateFont.setCharSet(Font.DEFAULT_CHARSET);
            dateFont.setColor(IndexedColors.BLACK.getIndex());
        }

        /**
         * @Description: 初始化表头行字体
         */
        private void initHeadFont() {
            //headFont.setFontName("宋体");
            headFont.setFontHeightInPoints((short) 10);
            headFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
            headFont.setCharSet(Font.DEFAULT_CHARSET);
            headFont.setColor(IndexedColors.BLACK.getIndex());
        }

        /**
         * @Description: 初始化内容行字体
         */
        private void initContentFont() {
            //contentFont.setFontName("宋体");
            contentFont.setFontHeightInPoints((short) 10);
            contentFont.setBoldweight(Font.BOLDWEIGHT_NORMAL);
            contentFont.setCharSet(Font.DEFAULT_CHARSET);
            contentFont.setColor(IndexedColors.BLACK.getIndex());
        }

        private void dispose() {
            if (this.wb != null) {
                this.wb.dispose();
            }
        }
    }
}