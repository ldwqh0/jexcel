package org.xyyh.jexcel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Target;

@Target({ ElementType.FIELD, ElementType.METHOD })
public @interface Col {
	String name() default "";
}
