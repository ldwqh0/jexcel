package org.xyyh.jexcel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Target;

@Target({ ElementType.FIELD, ElementType.METHOD })
public @interface Col {

	/**
	 * 规定某列的表头名称
	 * 
	 * @return
	 */
	String name() default "";

	/**
	 * 规定该列的排序
	 * 
	 * @return
	 */
	int sort() default 0;

}
