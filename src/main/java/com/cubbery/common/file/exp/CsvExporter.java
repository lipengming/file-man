/**
 * Copyright (c) 2016, www.cubbery.com. All rights reserved.
 */
package com.cubbery.common.file.exp;


import com.cubbery.common.file.Exporter;
import com.cubbery.common.file.FileInfo;
import com.cubbery.common.file.Header;

import java.io.IOException;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Map;

/**
 * <b>类描述</b>：   <br>
 * <b>创建人</b>：   <a href="mailto:cubber.zh@gmail,com">百墨</a> <br>
 * <b>创建时间</b>： 2016/5/26 - 17:04  <br>
 *
 * @version 1.0.0   <br>
 */
public class CsvExporter implements Exporter {
    /** 列分隔符 */
    private static final String CSV_CO_SEPARATOR = ",";
    /** 行分隔符 */
    private static final String CSV_RN_SEPARATOR = "\r\n";

    @Override
    public void export(FileInfo data) throws IOException {
        StringBuilder sb = new StringBuilder();
        //抬头行
        for(int a = 0; a < data.getHeaderIndex(); a++) {
            sb.append(CSV_RN_SEPARATOR);
        }
        //标题行
        for(Header h : data.getHeaders()) {
            sb.append(h.getColumnName()).append(CSV_CO_SEPARATOR);
        }
        sb.append(CSV_RN_SEPARATOR);
        for(Map<String,String> item : data.getData()) {
            for (Header header : data.getHeaders()) {
                String value = dateToStr(item.get(header.getFieldCode()));
                sb.append(value).append(CSV_CO_SEPARATOR);
            }
            sb.append(CSV_RN_SEPARATOR);
        }
        OutputStream out = data.getOutputStream();
        try {
            out.write(sb.toString().getBytes("UTF-8"));
        } finally {
            if(out != null) {
                out.flush();
                out.close();
            }
        }
    }

    public static String dateToStr(Object obj) {
        if(obj == null) {
            return "";
        }
        if(obj instanceof java.util.Date) {
            DateFormat df = new SimpleDateFormat(dayPatterns);
            return df.format((java.util.Date)obj);
        }
        return obj.toString();
    }

    public final static String dayPatterns = "yyyy-MM-dd";
}
