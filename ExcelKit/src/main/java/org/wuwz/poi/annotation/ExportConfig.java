package org.wuwz.poi.annotation;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by lifeng on 2018/6/1.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({java.lang.annotation.ElementType.FIELD})
public @interface ExportConfig {
    String value() default "field";

    short width() default -1;

    String convert() default "";

    short color() default 8;

    String replace() default "";

    String range() default "";

    //excle单元格数据类型,目前只支持货币和文本类型，为空就是文本，值不为空则为double类型
    String dataType() default "";
}
