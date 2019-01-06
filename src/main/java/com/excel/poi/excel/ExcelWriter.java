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
/**
 * @author NingWei
 */
package com.excel.poi.excel;

import static com.excel.poi.common.Constant.CHINESES_ATUO_SIZE_COLUMN_WIDTH_MAX;
import static com.excel.poi.common.Constant.CHINESES_ATUO_SIZE_COLUMN_WIDTH_MIN;
import static com.excel.poi.common.Constant.MAX_RECORD_COUNT_PEER_SHEET;
import static com.excel.poi.common.DateFormatUtil.format;
import com.excel.poi.entity.ExcelEntity;
import com.excel.poi.entity.ExcelPropertyEntity;
import com.excel.poi.function.ExportFunction;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * 导出具体实现类
 *
 * @author NingWei
 */
@Slf4j
public class ExcelWriter {

    private Integer rowAccessWindowSize;
    private ExcelEntity excelEntity;
    private Integer pageSize;
    private Integer recordCountPerSheet;
    private XSSFCellStyle headCellStyle;
    private Map<Integer, Integer> columnWidthMap = new HashMap<Integer, Integer>();


    public ExcelWriter(ExcelEntity excelEntity, Integer pageSize, Integer rowAccessWindowSize, Integer RecordCountPerSheet) {
        this.excelEntity = excelEntity;
        this.pageSize = pageSize;
        this.rowAccessWindowSize = rowAccessWindowSize;
        this.recordCountPerSheet = RecordCountPerSheet;
    }

    /**
     * @param param
     * @param exportFunction
     * @param <P>
     * @param <T>
     * @return
     * @throws InvocationTargetException
     * @throws NoSuchMethodException
     * @throws ParseException
     * @throws IllegalAccessException
     */
    public <P, T> SXSSFWorkbook generateWorkbook(P param, ExportFunction<P, T> exportFunction) throws Exception {
        SXSSFWorkbook workbook = new SXSSFWorkbook(rowAccessWindowSize);
        int sheetNo = 1;
        int rowNum = 1;
        List<ExcelPropertyEntity> propertyList = excelEntity.getPropertyList();
        //初始化第一行
        SXSSFSheet sheet = generateHeader(workbook, propertyList, excelEntity.getFileName());

        //生成其他行
        int firstPageNo = 1;
        while (true) {
            List<T> data = exportFunction.pageQuery(param, firstPageNo, pageSize);
            if (data == null || data.isEmpty()) {
                if (rowNum != 1) {
                    sizeColumWidth(sheet, propertyList.size());
                }
                log.warn("查询结果为空,结束查询!");
                break;
            }
            int dataSize = data.size();
            for (int i = 1; i <= dataSize; i++, rowNum++) {
                T queryResult = data.get(i - 1);
                Object convertResult = exportFunction.convert(queryResult);
                if (rowNum > MAX_RECORD_COUNT_PEER_SHEET) {
                    sizeColumWidth(sheet, propertyList.size());
                    sheet = generateHeader(workbook, propertyList, excelEntity.getFileName() + "_" + sheetNo);
                    sheetNo++;
                    rowNum = 1;
                    columnWidthMap.clear();
                }
                SXSSFRow row = sheet.createRow(rowNum);
                for (int j = 0; j < propertyList.size(); j++) {
                    SXSSFCell cell = row.createCell(j);
                    buildCellValue(cell, convertResult, propertyList.get(j));
                    calculateColumWidth(cell, j);
                }
            }
            if (data.size() < pageSize) {
                sizeColumWidth(sheet, propertyList.size());
                log.warn("查询结果数量小于pageSize,结束查询!");
                break;
            }
            firstPageNo++;
        }
        return workbook;
    }


    /**
     * 构建模板Excel
     *
     * @param <R>
     * @param <T>
     * @return
     */
    public SXSSFWorkbook generateTemplateWorkbook() {
        SXSSFWorkbook workbook = new SXSSFWorkbook(rowAccessWindowSize);

        List<ExcelPropertyEntity> propertyList = excelEntity.getPropertyList();
        SXSSFSheet sheet = generateHeader(workbook, propertyList, excelEntity.getFileName());

        SXSSFRow row = sheet.createRow(1);
        for (int j = 0; j < propertyList.size(); j++) {
            SXSSFCell cell = row.createCell(j);
            cell.setCellValue(propertyList.get(j).getTemplateCellValue());
            calculateColumWidth(cell, j);
        }
        sizeColumWidth(sheet, propertyList.size());
        return workbook;
    }

    /**
     * 构建多Sheet Excel
     *
     * @param param
     * @param exportFunction
     * @param <R>
     * @param <T>
     * @return
     * @throws InvocationTargetException
     * @throws NoSuchMethodException
     * @throws ParseException
     * @throws IllegalAccessException
     */
    public <R, T> SXSSFWorkbook generateMultiSheetWorkbook(R param, ExportFunction<R, T> exportFunction) throws Exception {
        int pageNo = 1;
        int sheetNo = 1;
        int rowNum = 1;
        SXSSFWorkbook workbook = new SXSSFWorkbook(rowAccessWindowSize);
        List<ExcelPropertyEntity> propertyList = excelEntity.getPropertyList();
        SXSSFSheet sheet = generateHeader(workbook, propertyList, excelEntity.getFileName());

        while (true) {
            List<T> data = exportFunction.pageQuery(param, pageNo, pageSize);
            if (data == null || data.isEmpty()) {
                if (rowNum != 1) {
                    sizeColumWidth(sheet, propertyList.size());
                }
                log.warn("查询结果为空,结束查询!");
                break;
            }
            for (int i = 1; i <= data.size(); i++, rowNum++) {
                T queryResult = data.get(i - 1);
                Object convertResult = exportFunction.convert(queryResult);
                if (rowNum > recordCountPerSheet) {
                    sizeColumWidth(sheet, propertyList.size());
                    sheet = generateHeader(workbook, propertyList, excelEntity.getFileName() + "_" + sheetNo);
                    sheetNo++;
                    rowNum = 1;
                    columnWidthMap.clear();
                }
                SXSSFRow bodyRow = sheet.createRow(rowNum);
                for (int j = 0; j < propertyList.size(); j++) {
                    SXSSFCell cell = bodyRow.createCell(j);
                    buildCellValue(cell, convertResult, propertyList.get(j));
                    calculateColumWidth(cell, j);
                }
            }
            if (data.size() < pageSize) {
                sizeColumWidth(sheet, propertyList.size());
                log.warn("查询结果数量小于pageSize,结束查询!");
                break;
            }
            pageNo++;
        }
        return workbook;
    }

    /**
     * 自动适配中文单元格
     *
     * @param sheet
     * @param cell
     * @param columnIndex
     */
    private void sizeColumWidth(SXSSFSheet sheet, Integer columnSize) {
        for (int j = 0; j < columnSize; j++) {
            if (columnWidthMap.get(j) != null) {
                sheet.setColumnWidth(j, columnWidthMap.get(j) * 256);
            }
        }
    }

    /**
     * 自动适配中文单元格
     *
     * @param sheet
     * @param cell
     * @param columnIndex
     */
    private void calculateColumWidth(SXSSFCell cell, Integer columnIndex) {
        int length = cell.getStringCellValue().getBytes().length;
        length = Math.max(length, CHINESES_ATUO_SIZE_COLUMN_WIDTH_MIN);
        length = Math.min(length, CHINESES_ATUO_SIZE_COLUMN_WIDTH_MAX);
        if (columnWidthMap.get(columnIndex) == null || columnWidthMap.get(columnIndex) < length) {
            columnWidthMap.put(columnIndex, length);
        }
    }

    /**
     * 初始化第一行的属性
     *
     * @param workbook
     * @param propertyList
     * @param sheetName
     * @return
     */
    private SXSSFSheet generateHeader(SXSSFWorkbook workbook, List<ExcelPropertyEntity> propertyList, String sheetName) {
        SXSSFSheet sheet = workbook.createSheet(sheetName);
        SXSSFRow headerRow = sheet.createRow(0);
        headerRow.setHeight((short) 600);
        CellStyle headCellStyle = getHeaderCellStyle(workbook);
        for (int i = 0; i < propertyList.size(); i++) {
            SXSSFCell cell = headerRow.createCell(i);
            cell.setCellStyle(headCellStyle);
            cell.setCellValue(propertyList.get(i).getColumnName());
            calculateColumWidth(cell, i);
        }
        return sheet;
    }

    /**
     * 构造 除第一行以外的其他行的列值
     *
     * @param cell
     * @param entity
     * @param property
     */
    private void buildCellValue(SXSSFCell cell, Object entity, ExcelPropertyEntity property) throws Exception {
        Field field = property.getFieldEntity();
        Object cellValue = field.get(entity);

        if (cellValue == null) {
            cell.setCellValue("");
        } else if (cellValue instanceof BigDecimal) {
            if (-1 == property.getScale()) {
                cell.setCellValue(cellValue.toString());
            } else {
                cell.setCellValue((((BigDecimal) cellValue).setScale(property.getScale(), property.getRoundingMode())).toString());
            }
        } else if (cellValue instanceof Date) {
            cell.setCellValue(format(property.getDateFormat(), (Date) cellValue));
        } else {
            cell.setCellValue(cellValue.toString());
        }
    }


    public CellStyle getHeaderCellStyle(SXSSFWorkbook workbook) {
        if (headCellStyle == null) {
            headCellStyle = workbook.getXSSFWorkbook().createCellStyle();
            headCellStyle.setBorderTop(BorderStyle.NONE);
            headCellStyle.setBorderRight(BorderStyle.NONE);
            headCellStyle.setBorderBottom(BorderStyle.NONE);
            headCellStyle.setBorderLeft(BorderStyle.NONE);
            headCellStyle.setAlignment(HorizontalAlignment.CENTER);// 居中
            headCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 居中
            XSSFColor color = new XSSFColor(new java.awt.Color(217, 217, 217));
            headCellStyle.setFillForegroundColor(color);
            headCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Font font = workbook.createFont();
            font.setFontName("微软雅黑");
            font.setColor(IndexedColors.ROYAL_BLUE.index);
            font.setBold(true);
            headCellStyle.setFont(font);
            headCellStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));
        }
        return headCellStyle;
    }
}
