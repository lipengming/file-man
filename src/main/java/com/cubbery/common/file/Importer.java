/**
 * Copyright (c) 2016, www.cubbery.com. All rights reserved.
 */
package com.cubbery.common.file;

import java.io.IOException;

/**
 * <b>类描述</b>：   <br>
 * <b>创建人</b>：   <a href="mailto:cubber.zh@gmail,com">百墨</a> <br>
 * <b>创建时间</b>： 2016/5/25 - 17:43  <br>
 *
 * @version 1.0.0   <br>
 */
public interface Importer {
    /**
     * 导入：账单文件解析</br>
     *
     * @param fileInfo  文件信息，必填项：</br>
     *                  <ul>
     *                      <li>name:文件名(excel格式为必填)</li>
     *                      <li>headerIndex:标题行</li>
     *                      <li>sheetIndex:表单编号(excel格式为必填)</li>
     *                      <li>inputStream:输入流</li>
     *                  </ul>
     *
     * @return 对于解析结果，直接包装在参数的fileInfo里面。非异常情况下：
     *                  <ul>
     *                      <li>headers:标题信息</li>
     *                      <li>data:数据信息</li>
     *                  </ul>
     *
     * @exception java.io.IOException 流读取失败
     * @exception java.lang.RuntimeException    文件格式校验失败
     */
    void resolve(FileInfo fileInfo, Converter<String, String> converter) throws IOException;
}
