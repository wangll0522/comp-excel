package com.wangll.comp.excel.utils.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Description: Excel
 * @package: com.beyond.transafemrg.common.utils.anno.
 * Created by ll_wang on 16/4/1.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSupport {
    //字段名
    String name() default "";
    //导出的时候是否使用,false时过滤字段
    boolean use() default true;
    //是否为编码
    boolean code() default false;
    //列宽
    short cellWidth() default 12;
    //自动换行
    boolean wrap() default false;
    //排序
    int sort() default 100;
    //日期格式化
    String format() default "yyyy-MM-dd HH:mm:ss";
}
