/**
 * Copyright (c) 2016, www.cubbery.com. All rights reserved.
 */
package com.cubbery.common.file.imp;

import com.cubbery.common.file.Converter;
import com.cubbery.common.file.FileInfo;
import com.cubbery.common.file.Header;
import com.cubbery.common.file.Importer;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.*;
import java.util.logging.Logger;

/**
 * <b>类描述</b>：   <br>
 * <b>创建人</b>：   <a href="mailto:cubber.zh@gmail,com">百墨</a> <br>
 * <b>创建时间</b>： 2016/5/26 - 15:46  <br>
 *
 * @version 1.0.0   <br>
 */
public class ExcelImporter implements Importer {
    private final static Logger LOG = Logger.getLogger(ExcelImporter.class.getName());

    @Override
    public void resolve(FileInfo fileInfo,Converter<String,String> converter) throws IOException {
        ExcelHandler excelHandler = new ExcelHandler();
        //1、初始化信息，参数校验
        excelHandler.init(fileInfo);
        //2、解析表头，放到header里面
        excelHandler.resolveHeader(fileInfo,converter);
        //3、解析内容
        excelHandler.resolveData(fileInfo);
    }

    class ExcelHandler {
        //标题行号
        private int headerIndex;
        //表单号
        private int sheetIndex;
        //工作薄
        private Workbook wb;
        //工作表单
        private Sheet sheet;

        private void resolveData(FileInfo fileInfo) {
            int dataRow = headerIndex + 1;
            int lastRowNum = sheet.getLastRowNum();
            List<Map<String,String>> dataList = new ArrayList<Map<String, String>>(lastRowNum - headerIndex);
            int columnSize = fileInfo.getHeaders().size();
            for (int i = dataRow; i <= lastRowNum; i++) {
                Row row = this.getRow(i);
                //没有值的行数
                int blankColumns = 0;
                Map<String,String> data = new HashMap<String, String>(columnSize);
                for(int index : fileInfo.getMap().keySet()) {
                    String cellValue = getCellValue(row, index);
                    if(StringUtils.isBlank(cellValue)) {
                        blankColumns++;
                    }
                    String fieldCode = fileInfo.getMap().get(index);
                    data.put(fieldCode, cellValue);
                }
                if(blankColumns == columnSize) {
                    continue;//改行数据不可用
                }
                dataList.add(data);
            }
            fileInfo.setData(dataList);
        }

        private void resolveHeader(FileInfo fileInfo,Converter<String,String> converter) {
            Row headerRow = getRow(headerIndex);
            Iterator<Cell> iterator = headerRow.iterator();
            List<Header> headers = new ArrayList<Header>();
            while (iterator.hasNext()) {
                Cell cell = iterator.next();
                Object cellValue = getCellValue(cell);
                if(cellValue == null) {
                    break;
                }
                Header h = new Header();
                h.setColumnName(cellValue.toString());
                h.setColumnIndex(cell.getColumnIndex());
                String fieldCode = converter.convert(cellValue.toString());
                if(StringUtils.isBlank(fieldCode)) {
                    //存在无法识别的列
                    LOG.info("存在无法识别的列: " + cellValue.toString());
                    throw new RuntimeException("存在无法识别的列: " + cellValue.toString());
                }
                if(fileInfo.getKey() != null && fieldCode.contains(fileInfo.getKey())) {
                    fileInfo.addColumn(fieldCode);
                }
                h.setFieldCode(fieldCode);
                headers.add(h);
            }
            fileInfo.setHeaders(headers);
        }

        private void init(FileInfo fileInfo) throws IOException {
            String fileName = fileInfo.getName();
            if (StringUtils.isBlank(fileName)) {
                LOG.info("导入文档名称不能为空!");
                throw new RuntimeException("导入文档名称不能为空!");
            } else if (fileName.toLowerCase().endsWith("xls")) {
                this.wb = new HSSFWorkbook(fileInfo.getInputStream());
            } else if (fileName.toLowerCase().endsWith("xlsx")) {
                this.wb = new XSSFWorkbook(fileInfo.getInputStream());
            } else {
                LOG.info("文档格式不是Excel格式! 文件名：" + fileName);
                throw new RuntimeException("文档格式不是Excel格式! 文件名：" + fileName);
            }
            this.headerIndex = fileInfo.getHeaderIndex();
            this.sheetIndex = fileInfo.getSheetIndex();
            this.sheet = this.wb.getSheetAt(sheetIndex);
            fileInfo.setSheetName(this.sheet.getSheetName());
        }

        public Row getRow(int rowNum) {
            return this.sheet.getRow(rowNum);
        }

        public String getCellValue(Cell cell) {
            String val = null;
            if (cell != null) {
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        double v = cell.getNumericCellValue();
                        val = new BigDecimal(v).toString();
                        break;
                    case Cell.CELL_TYPE_STRING:
                        val = cell.getStringCellValue(); break;
                    case Cell.CELL_TYPE_FORMULA:
                        val = cell.getCellFormula(); break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        Boolean b = cell.getBooleanCellValue();
                        val = b.toString();
                        break;
                    case Cell.CELL_TYPE_ERROR:
                        Byte by = cell.getErrorCellValue();
                        val = by.toString();
                        break;
                }
            }
            return val;
        }

        public String getCellValue(Row row, int column){
            String val = null;
            try{
                Cell cell = row.getCell(column);
                val = getCellValue(cell);
            }catch (Exception e) {
                //ignore
            }
            return val;
        }
    }

}