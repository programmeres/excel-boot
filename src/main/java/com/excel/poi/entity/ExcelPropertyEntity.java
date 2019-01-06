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
package com.excel.poi.entity;

import java.lang.reflect.Field;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

/**
 * @author NingWei
 */
@Getter
@Setter
@Builder
public class ExcelPropertyEntity {

    /**
     * excelModel字段Field
     */
    private Field fieldEntity;
    /**
     * excel列名称
     *
     * @return
     */
    private String columnName;
    /**
     * 默认单元格值
     *
     * @return
     */
    private String templateCellValue;
    /**
     * 日期格式 默认 yyyy-MM-dd HH:mm:ss
     *
     * @return
     */
    private String dateFormat;
    /**
     * 正则表达式校验
     *
     * @return
     */
    private String regex;
    /**
     * 正则表达式校验失败返回的错误信息,regex配置后生效
     *
     * @return
     */
    private String regexMessage;
    /**
     * BigDecimal精度 默认:2
     *
     * @return
     */
    private Integer scale;
    /**
     * BigDecimal 舍入规则 默认:2
     *
     * @return
     */
    private Integer roundingMode;
    /**
     * @return 是否必填
     */
    private Boolean required;
}
