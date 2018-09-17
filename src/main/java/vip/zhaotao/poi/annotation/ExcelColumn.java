package vip.zhaotao.poi.annotation;

import java.lang.annotation.*;

@Documented
@Target(value = {ElementType.FIELD})
@Retention(value = RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

    /**
     * Column name
     */
    String name();

    /**
     * Column number (0-based)
     */
    int number() default 0;

    /**
     * Column data format
     */
    String format() default "";
}
