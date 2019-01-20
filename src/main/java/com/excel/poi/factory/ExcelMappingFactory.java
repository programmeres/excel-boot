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
package com.excel.poi.factory;

import com.excel.poi.annotation.ExportField;
import com.excel.poi.annotation.ImportField;
import com.excel.poi.entity.ExcelEntity;
import com.excel.poi.entity.ExcelPropertyEntity;
import com.excel.poi.exception.ExcelBootException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * @author NingWei
 */
public class ExcelMappingFactory {

    /**
     * 根据指定Excel实体获取导入Excel文件相关信息
     *
     * @param clazz
     * @return
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public static ExcelEntity loadImportExcelClass(Class clazz) {
        List<ExcelPropertyEntity> propertyList = new ArrayList<ExcelPropertyEntity>();

        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            ImportField importField = field.getAnnotation(ImportField.class);
            if (null != importField) {
                field.setAccessible(true);
                ExcelPropertyEntity excelPropertyEntity = ExcelPropertyEntity.builder()
                        .fieldEntity(field)
                        .required(importField.required())
                        .dateFormat(importField.dateFormat().trim())
                        .regex(importField.regex().trim())
                        .regexMessage(importField.regexMessage().trim())
                        .scale(importField.scale())
                        .roundingMode(importField.roundingMode())
                        .build();
                propertyList.add(excelPropertyEntity);
            }
        }
        if (propertyList.isEmpty()) {
            throw new ExcelBootException("[{}] 类未找到标注@ImportField注解的属性!", clazz.getName());
        }
        ExcelEntity excelMapping = new ExcelEntity();
        excelMapping.setPropertyList(propertyList);
        return excelMapping;

    }

    /**
     * 根据指定Excel实体获取导出Excel文件相关信息
     *
     * @param clazz
     * @return
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public static ExcelEntity loadExportExcelClass(Class<?> clazz, String fileName) {
        List<ExcelPropertyEntity> propertyList = new ArrayList<ExcelPropertyEntity>();

        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            ExportField exportField = field.getAnnotation(ExportField.class);
            if (null != exportField) {
                field.setAccessible(true);
                ExcelPropertyEntity excelPropertyEntity = ExcelPropertyEntity.builder()
                        .fieldEntity(field)
                        .columnName(exportField.columnName().trim())
                        .scale(exportField.scale())
                        .roundingMode(exportField.roundingMode())
                        .dateFormat(exportField.dateFormat().trim())
                        .templateCellValue(exportField.defaultCellValue().trim())
                        .build();
                propertyList.add(excelPropertyEntity);
            }
        }
        if (propertyList.isEmpty()) {
            throw new ExcelBootException("[{}]类未找到标注@ExportField注解的属性!", clazz.getName());
        }
        ExcelEntity excelMapping = new ExcelEntity();
        excelMapping.setPropertyList(propertyList);
        excelMapping.setFileName(fileName);
        return excelMapping;
    }


}
