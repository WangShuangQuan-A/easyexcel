package com.wsq.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import static com.wsq.excel.metadata.ExcelColumnProperty.DEFAULT_INDEX;

/**
 * @author jipengfei
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelProperty {

     /**
      * @return
      */
     String value() default "";


     /**
      * @return
      */
     int index() default DEFAULT_INDEX;

     /**
      *
      * default @see com.alibaba.excel.util.TypeUtil
      * if default is not  meet you can set format
      *
      * @return
      */
     String format() default "";
}
