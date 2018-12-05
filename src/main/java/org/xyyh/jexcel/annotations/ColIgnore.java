package org.xyyh.jexcel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Target;

/**
 * 标注一个类的某个字段或者方法应被忽略
 * 
 * @author LiDong
 *
 */
@Target({ ElementType.FIELD, ElementType.METHOD })
public @interface ColIgnore {

}
