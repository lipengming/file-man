/**
 * Copyright (c) 2016, www.cubbery.com. All rights reserved.
 */
package com.cubbery.common.file;


import com.cubbery.common.file.exp.CsvExporter;
import com.cubbery.common.file.exp.ExcelExporter;
import com.cubbery.common.file.imp.CsvImporter;
import com.cubbery.common.file.imp.ExcelImporter;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;

/**
 * <b>类描述</b>：   <br>
 * <b>创建人</b>：   <a href="mailto:cubber.zh@gmail,com">百墨</a> <br>
 * <b>创建时间</b>： 2016/6/1 - 15:40  <br>
 *
 * @version 1.0.0   <br>
 */
public enum FileMan {
    CSV(new CsvImporter(),new CsvExporter(), Arrays.asList("csv")) {
        @Override
        boolean check(String suffix) {
            if(suffix == null) {
                return false;
            }
            String suf = suffix.toLowerCase();
            return getSuffix().contains(suf);
        }
    },
    EXCEL(new ExcelImporter(),new ExcelExporter(),Arrays.asList("xls","xlsx")) {
        @Override
        boolean check(String suffix) {
            if(suffix == null) {
                return false;
            }
            String suf = suffix.toLowerCase();
            return getSuffix().contains(suf);
        }
    };

    private Importer importer;
    private Exporter exporter;
    private List<String> suffix;

    FileMan(Importer importer,Exporter exporter, List<String> suffix) {
        this.importer = importer;
        this.exporter = exporter;
        this.suffix = suffix;
    }

    private static FileMan getFileType(String fileName) {
        String suffix = getSuffix(fileName);
        for(FileMan f : values()) {
            if(f.check(suffix)) {
                return f;
            }
        }
        return null;
    }

    public static void resolve(FileInfo fileInfo, Converter<String, String> converter) throws IOException {
        FileMan man = getFileType(fileInfo.getName());
        if(man == null) {
            throw new RuntimeException("格式错误！");
        }
        Importer importer = man.getImporter();
        importer.resolve(fileInfo, converter);
    }

    public static void export(FileInfo fileInfo) throws IOException {
        FileMan man = getFileType(fileInfo.getName());
        if(man == null) {
            throw new RuntimeException("格式错误！");
        }
        Exporter exporter = man.getExporter();
        exporter.export(fileInfo);
    }

    public Importer getImporter() {
        return importer;
    }

    public Exporter getExporter() {
        return exporter;
    }

    public List<String> getSuffix() {
        return suffix;
    }

    abstract boolean check(String suffix);

    public static String getPrefix(String filePath) {
        if(filePath == null || filePath.length() < 2) {
            return filePath;
        }
        String prefix = filePath.substring(0,filePath.lastIndexOf(".") - 1);
        return prefix;
    }

    public static String getSuffix(String filePath) {
        if(filePath == null || filePath.length() < 2) {
            return filePath;
        }
        String suffix = filePath.substring(filePath.lastIndexOf(".") + 1);
        return suffix;
    }

}
