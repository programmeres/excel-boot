/*
 * Copyright 2018 NingWei (ningww1@126.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * </p>
 */
package com.excel.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.math.BigDecimal;

/**
 * @author NingWei
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ImportField {

    /**
     * @return 是否必填
     */
    boolean required() default false;

    /**
     * 日期格式 默认 yyyy-MM-dd HH:mm:ss
     *
     * @return
     */
    String dateFormat() default "yyyy-MM-dd HH:mm:ss";

    /**
     * 正则表达式校验
     *
     * @return
     */
    String regex() default "";

    /**
     * 正则表达式校验失败返回的错误信息,regex配置后生效
     *
     * @return
     */
    String regexMessage() default "正则表达式验证失败";

    /**
     * BigDecimal精度 默认:2
     *
     * @return
     */
    int scale() default 2;

    /**
     * BigDecimal 舍入规则 默认:BigDecimal.ROUND_HALF_EVEN
     *
     * @return
     */
    int roundingMode() default BigDecimal.ROUND_HALF_EVEN;
}
