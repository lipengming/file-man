/**
 * Copyright (c) 2016, www.cubbery.com. All rights reserved.
 */
package com.cubbery.common.file;

/**
 * <b>类描述</b>：   <br>
 * <b>创建人</b>：   <a href="mailto:cubber.zh@gmail,com">百墨</a> <br>
 * <b>创建时间</b>： 2016/5/26 - 17:46  <br>
 *
 * @version 1.0.0   <br>
 */
public interface Converter<R,T> {
    /**
     * 一个值转换成另外一个值
     *
     * @param one       for example: column_name
     * @return          return: field_code
     */
    T convert(R one);
}
