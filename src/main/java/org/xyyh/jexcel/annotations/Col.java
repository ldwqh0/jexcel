package org.xyyh.jexcel.annotations;

import java.lang.annotation.*;

@Target({ ElementType.FIELD, ElementType.METHOD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
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
	int sort() default -1;

}
