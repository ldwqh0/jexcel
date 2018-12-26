package org.xyyh.jexcel.annotations;

import java.lang.annotation.*;

/**
 * 标注一个类的某个字段或者方法应被忽略
 * 
 * @author LiDong
 *
 */
@Target({ ElementType.FIELD, ElementType.METHOD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ColIgnore {

    /**
     * 默认忽略此表头
     * @return
     */
    boolean value() default true;
}
