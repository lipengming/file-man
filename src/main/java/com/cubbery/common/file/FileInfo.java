/**
 * Copyright (c) 2016, www.cubbery.com. All rights reserved.
 */
package com.cubbery.common.file;

import java.io.InputStream;
import java.io.OutputStream;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * <b>类描述</b>：   <br>
 * <b>创建人</b>：   <a href="mailto:cubber.zh@gmail,com">百墨</a> <br>
 * <b>创建时间</b>： 2016/5/26 - 15:48  <br>
 *
 * @version 1.0.0   <br>
 */
public class FileInfo implements Serializable {
    //标题
    private String title;
    //文件名称
    private String name;
    //表头行号
    private int headerIndex;
    //表单号
    private int sheetIndex;
    //表单名称
    private String sheetName;
    //表头
    private transient List<Header> headers;
    //数据
    private transient List<Map<String,String>> data;
    //输出流
    private transient OutputStream outputStream;
    //输入流
    private transient InputStream inputStream;
    //映射关系key:列号，value:fieldCode
    private transient Map<Integer,String> map;
    //关键信息(用于区分是复合账单还是单一类型的账单)
    private KeyPoint<String> keyPoint;

    public FileInfo() {
        keyPoint = new KeyPoint<String>();
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    public InputStream getInputStream() {
        return inputStream;
    }

    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    public List<Header> getHeaders() {
        return headers;
    }

    public void setHeaders(List<Header> headers) {
        this.headers = headers;
        if(headers == null || headers.size() < 1) {
            return;
        }
        this.map = new HashMap<Integer, String>(headers.size());
        for(Header header : headers) {
            this.map.put(header.getColumnIndex(),header.getFieldCode());
        }
    }

    public List<Map<String, String>> getData() {
        return data;
    }

    public void setData(List<Map<String, String>> data) {
        this.data = data;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getHeaderIndex() {
        return headerIndex;
    }

    public void setHeaderIndex(int headerIndex) {
        this.headerIndex = headerIndex;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public Map<Integer, String> getMap() {
        return map;
    }

    public List<String> getColumns() {
        return keyPoint.getColumns();
    }

    public void addColumn(String column) {
        this.keyPoint.getColumns().add(column);
    }

    public void setKey(String key) {
        this.keyPoint.setKey(key);
        this.keyPoint.setColumns(new ArrayList<String>());
    }

    public String getKey() {
        return this.keyPoint.getKey();
    }

    public void destroy() {
        this.headers = null;
        this.data = null;
        this.map = null;
        this.inputStream = null;
        this.outputStream = null;
    }

    public class KeyPoint<K> {
        K key;//关键信息编码，如根据费用列来区分是复合还是单一，那么设置为（fee）, 匹配到fee相关的列，放入到cloumns里面。
        List<K> columns;//所有列

        public K getKey() {
            return key;
        }

        public void setKey(K key) {
            this.key = key;
        }

        public List<K> getColumns() {
            return columns;
        }

        public void setColumns(List<K> columns) {
            this.columns = columns;
        }
    }
}
