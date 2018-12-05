package org.xyyh.jexcel.annotations;

/**
 * 标注将一个对象转换为excel时，一个excel sheet的相关属性
 * 
 * @author LiDong
 *
 */
public @interface Sheet {
	/**
	 * 规定sheet的名称
	 * 
	 * @return
	 */
	String name() default "";
}
