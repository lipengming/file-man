/**
 * Copyright (c) 2016, www.cubbery.com. All rights reserved.
 */
package com.cubbery.common.file;

import java.io.IOException;

/**
 * <b>类描述</b>：   <br>
 * <b>创建人</b>：   <a href="mailto:cubber.zh@gmail,com">百墨</a> <br>
 * <b>创建时间</b>： 2016/5/26 - 15:30  <br>
 *
 * @version 1.0.0   <br>
 */
public interface Exporter {
    /**
     * 导出：账单文件导出
     *
     * @param data     文件信息。</br>
     *                 <ul>
     *                      <li>headers:标题信息必须<b>按照columnIndex升序排序</b></li>
     *                      <li>data:里面的map数据必须和headers的长度保持一致</li>
     *                      <li>title：标题（可填）</li>
     *                      <li>name：文件名（可填）</li>
     *                      <li>headerIndex：标题列编号（可填）</li>
     *                      <li>sheetName：表单名称（可填）</li>
     *                      <li>outputStream：输出流</li>
     *                 </ul>
     *
     * @return 直接写出到流。不返回任何数据。
     *
     * @exception java.io.IOException         文件操作异常
     * @exception RuntimeException            必填参数错误
     */
    void export(FileInfo data) throws IOException;
}
