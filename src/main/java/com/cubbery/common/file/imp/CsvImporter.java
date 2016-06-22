/**
 * Copyright (c) 2016, www.cubbery.com. All rights reserved.
 */
package com.cubbery.common.file.imp;


import com.cubbery.common.file.Converter;
import com.cubbery.common.file.FileInfo;
import com.cubbery.common.file.Header;
import com.cubbery.common.file.Importer;
import org.apache.any23.encoding.TikaEncodingDetector;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.util.*;
import java.util.logging.Logger;

/**
 * <b>类描述</b>：   <br>
 * <b>创建人</b>：   <a href="mailto:cubber.zh@gmail,com">百墨</a> <br>
 * <b>创建时间</b>： 2016/5/26 - 17:03  <br>
 *
 * @version 1.0.0   <br>
 */
public class CsvImporter implements Importer {
    private final static Logger LOG = Logger.getLogger(CsvImporter.class.getName());
    @Override
    public void resolve(FileInfo fileInfo, Converter<String,String> converter) throws IOException {
        InputStreamReader fr = null;
        BufferedReader br = null;
        try {
            BufferedInputStream bis = new BufferedInputStream(fileInfo.getInputStream());
            String charset = guessCharset(bis);
            fr = new InputStreamReader(bis,charset);
            br = new BufferedReader(fr);
            //先读取流，如果失败，不再创建集合
            List<Map<String, String>> dataList = new ArrayList<Map<String, String>>();
            String rec = null;// 一行
            for (int line = 0; (rec = br.readLine()) != null; line++) {
                if (line == fileInfo.getHeaderIndex()) {
                    resolveHeader(fileInfo, rec,converter);
                    continue;
                }
                if(fileInfo.getMap().isEmpty()) {
                    //头部信息还未读取
                    continue;
                }
                dataList.add(resolveContent(rec, fileInfo));
            }
            fileInfo.setData(dataList);
        } finally {
            //流关闭操作，但不包括fileInfo中的流关闭操作
            if (fr != null) {
                fr.close();
            }
            if (br != null) {
                br.close();
            }
        }
    }

    /**
     * 单行文件内容解析
     *
     * @param rec       单行文本
     * @param fileInfo  输入文件信息
     * @return          key:fieldCode|value为取值
     * @exception RuntimeException 文件格式异常
     */
    private Map<String, String> resolveContent(String rec, FileInfo fileInfo) throws RuntimeException {
        Map<String, String> data = new HashMap<String, String>();
        String[] lines = splitLine(rec);
        for (int a = 0; a < lines.length; a++) {
            String fieldCode = fileInfo.getMap().get(a);
            data.put(fieldCode,lines[a]);
        }
        return data;
    }

    /**
     * 标题行解析
     *
     * @param fileInfo      输入文件信息
     * @param rec           标题行
     * @param converter     转换器
     * @exception RuntimeException 文件格式异常
     */
    private void resolveHeader(FileInfo fileInfo, String rec, Converter<String,String> converter) {
        String[] headers = splitLine(rec);
        if (headers.length < 0) {
            throw new RuntimeException("Error!");
        }
        List<Header> hs = new ArrayList<Header>(headers.length);
        for (int a = 0; a < headers.length; a++) {
            if(StringUtils.isBlank(headers[a])) {
                continue;
            }
            Header header = new Header();
            header.setColumnIndex(a);
            header.setColumnName(headers[a]);
            String fieldCode = converter.convert(headers[a]);
            if(StringUtils.isBlank(fieldCode)) {
                //存在无法识别的列
                LOG.info("存在无法识别的列: " + headers[a]);
                throw new RuntimeException("存在无法识别的列: " + headers[a]);
            }
            if(fileInfo.getKey() != null && fieldCode.contains(fileInfo.getKey())) {
                fileInfo.addColumn(fieldCode);
            }
            header.setFieldCode(fieldCode);
            hs.add(header);
        }
        fileInfo.setHeaders(hs);
    }

    /**
     * 单行解析
     *
     * @param src                   字符串
     * @return                      解析每个column的值组成的数组
     * @throws RuntimeException     文件格式错误
     */
    private static String[] splitLine(String src) throws RuntimeException {
        if (src == null || src.equals("")) return new String[0];
        StringBuilder sb = new StringBuilder();
        Vector<String> result = new Vector<String>();
        boolean beginWithQuote = false;
        for (int i = 0; i < src.length(); i++) {
            char ch = src.charAt(i);
            if (ch == '\"') {
                if (beginWithQuote) {
                    i++;
                    if (i >= src.length()) {
                        result.addElement(sb.toString());
                        sb = new StringBuilder();
                        beginWithQuote = false;
                    } else {
                        ch = src.charAt(i);
                        if (ch == '\"') {
                            sb.append(ch);
                        } else if (ch == ',') {
                            result.addElement(sb.toString());
                            sb = new StringBuilder();
                            beginWithQuote = false;
                        } else {
                            LOG.info("CVS文件错误的使用了双引号！");
                            throw new RuntimeException("CVS文件错误的使用了双引号");
                        }
                    }
                } else if (sb.length() == 0) {
                    beginWithQuote = true;
                } else {
                    LOG.info("CVS文件错误的使用了双引号！");
                    throw new RuntimeException("CVS文件错误的使用了双引号");
                }
            } else if (ch == ',') {
                if (beginWithQuote) {
                    sb.append(ch);
                } else {
                    result.addElement(sb.toString());
                    sb = new StringBuilder();
                    beginWithQuote = false;
                }
            } else {
                sb.append(ch);
            }
        }
        if (sb.length() != 0) {
            if (beginWithQuote) {
                LOG.info("CVS文件错误的使用了双引号！");
                throw new RuntimeException("CVS文件错误的使用了双引号");
            } else {
                result.addElement(sb.toString());
            }
        }
        String rs[] = new String[result.size()];
        for (int i = 0; i < rs.length; i++) {
            rs[i] = result.elementAt(i);
        }
        return rs;
    }

    public static String guessCharset(InputStream is) throws IOException {
        return new TikaEncodingDetector().guessEncoding(is);
    }
}
