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
package com.excel.poi.excel;

import com.excel.poi.common.Constant;
import static com.excel.poi.common.DateFormatUtil.parse;
import static com.excel.poi.common.RegexUtil.isMatch;
import com.excel.poi.common.StringUtil;
import static com.excel.poi.common.StringUtil.convertNullTOZERO;
import com.excel.poi.entity.ErrorEntity;
import com.excel.poi.entity.ExcelEntity;
import com.excel.poi.entity.ExcelPropertyEntity;
import com.excel.poi.exception.AllEmptyRowException;
import com.excel.poi.exception.ExcelBootException;
import com.excel.poi.function.ImportFunction;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.ExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * @author NingWei
 */
@Slf4j
public class ExcelReader extends DefaultHandler {
    private Integer currentSheetIndex = -1;
    private Integer currentRowIndex = 0;
    private Integer excelCurrentCellIndex = 0;
    private ExcelCellType cellFormatStr;
    private String currentCellLocation;
    private String previousCellLocation;
    private String endCellLocation;
    private SharedStringsTable mSharedStringsTable;
    private String currentCellValue;
    private Boolean isNeedSharedStrings = false;
    private ExcelEntity excelMapping;
    private ImportFunction importFunction;
    private Class excelClass;
    private List<String> cellsOnRow = new ArrayList<String>();
    private Integer beginReadRowIndex;
    private Integer dataCurrentCellIndex = -1;


    public ExcelReader(Class entityClass,
                       ExcelEntity excelMapping,
                       ImportFunction importFunction) {
        this(entityClass, excelMapping, 1, importFunction);
    }

    public ExcelReader(Class entityClass,
                       ExcelEntity excelMapping,
                       Integer beginReadRowIndex,
                       ImportFunction importFunction) {
        this.excelClass = entityClass;
        this.excelMapping = excelMapping;
        this.beginReadRowIndex = beginReadRowIndex;
        this.importFunction = importFunction;
    }

    public void process(InputStream in)
            throws IOException, OpenXML4JException, SAXException {
        OPCPackage opcPackage = null;
        InputStream sheet = null;
        InputSource sheetSource;
        try {
            opcPackage = OPCPackage.open(in);
            XSSFReader xssfReader = new XSSFReader(opcPackage);
            XMLReader parser = this.fetchSheetParser(xssfReader.getSharedStringsTable());

            Iterator<InputStream> sheets = xssfReader.getSheetsData();
            while (sheets.hasNext()) {
                currentRowIndex = 0;
                currentSheetIndex++;
                try {
                    sheet = sheets.next();
                    sheetSource = new InputSource(sheet);
                    try {
                        log.info("开始读取第{}个Sheet!", currentSheetIndex + 1);
                        parser.parse(sheetSource);
                    } catch (AllEmptyRowException e) {
                        log.warn(e.getMessage());
                    } catch (Exception e) {
                        throw new ExcelBootException(e, "第{}个Sheet,第{}行,第{}列,系统发生异常! ", currentSheetIndex + 1, currentRowIndex + 1, dataCurrentCellIndex + 1);
                    }
                } finally {
                    if (sheet != null) {
                        sheet.close();
                    }
                }
            }
        } finally {
            if (opcPackage != null) {
                opcPackage.close();
            }
        }
    }

    /**
     * 获取sharedStrings.xml文件的XMLReader对象
     *
     * @param sst
     * @return
     * @throws SAXException
     */
    private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        this.mSharedStringsTable = sst;
        parser.setContentHandler(this);
        return parser;
    }

    /**
     * 开始读取一个标签元素
     *
     * @param uri
     * @param localName
     * @param name
     * @param attributes
     */
    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) {
        if (Constant.CELL.equals(name)) {
            String xyzLocation = attributes.getValue(Constant.XYZ_LOCATION);
            previousCellLocation = null == previousCellLocation ? xyzLocation : currentCellLocation;
            currentCellLocation = xyzLocation;
            String cellType = attributes.getValue(Constant.CELL_T_PROPERTY);
            isNeedSharedStrings = (null != cellType && cellType.equals(Constant.CELL_S_VALUE));
            setCellType(cellType);
        }
        currentCellValue = "";
    }


    /**
     * 加载v标签中间的值
     *
     * @param chars
     * @param start
     * @param length
     */
    @Override
    public void characters(char[] chars, int start, int length) {
        currentCellValue = currentCellValue.concat(new String(chars, start, length));
    }

    /**
     * 结束读取一个标签元素
     *
     * @param uri
     * @param localName
     * @param name
     * @throws SAXException
     */
    @Override
    public void endElement(String uri, String localName, String name) {
        if (Constant.CELL.equals(name)) {
            if (isNeedSharedStrings && !StringUtil.isBlank(currentCellValue) && StringUtil.isNumeric(currentCellValue)) {
                int index = Integer.parseInt(currentCellValue);
                currentCellValue = new XSSFRichTextString(mSharedStringsTable.getEntryAt(index)).toString();
            }
            if (!currentCellLocation.equals(previousCellLocation) && currentRowIndex != 0) {
                for (int i = 0; i < countNullCell(currentCellLocation, previousCellLocation); i++) {
                    cellsOnRow.add(excelCurrentCellIndex, "");
                    excelCurrentCellIndex++;
                }
            }
            if (currentRowIndex != 0 || !"".equals(currentCellValue.trim())) {
                String value = this.getCellValue(currentCellValue.trim());
                cellsOnRow.add(excelCurrentCellIndex, value);
                excelCurrentCellIndex++;
            }
        }
        else if (Constant.ROW.equals(name)) {
            if (currentRowIndex == 0) {
                endCellLocation = currentCellLocation;
                int propertySize = excelMapping.getPropertyList().size();
                if (cellsOnRow.size() != propertySize) {
                    throw new ExcelBootException("Excel有效列数不等于标注注解的属性数量!Excel列数:{},标注注解的属性数量:{}", cellsOnRow.size(), propertySize);
                }
            }
            if (null != endCellLocation) {
                for (int i = 0; i <= countNullCell(endCellLocation, currentCellLocation); i++) {
                    cellsOnRow.add(excelCurrentCellIndex, "");
                    excelCurrentCellIndex++;
                }
            }
            try {
                this.assembleData();
            } catch (AllEmptyRowException e) {
                throw e;
            } catch (Exception e) {
                throw new ExcelBootException(e);
            }
            cellsOnRow.clear();
            currentRowIndex++;
            dataCurrentCellIndex = -1;
            excelCurrentCellIndex = 0;
            previousCellLocation = null;
            currentCellLocation = null;
        }

    }

    /**
     * 根据c节点的t属性获取单元格格式
     * 根据c节点的s属性获取单元格样式,去styles.xml文件找相应样式
     *
     * @param cellType     xml中单元格格式属性
     * @param cellStyleStr xml中样式属性
     */
    private void setCellType(String cellType) {
        if ("inlineStr".equals(cellType)) {
            cellFormatStr = ExcelCellType.INLINESTR;
        } else if ("s".equals(cellType) || cellType == null) {
            cellFormatStr = ExcelCellType.STRING;
        } else {
            throw new ExcelBootException("Excel单元格格式未设置成文本或者常规!单元格格式:{}", cellType);
        }
    }

    /**
     * 根据数据类型获取数据
     *
     * @param value
     * @return
     */
    private String getCellValue(String value) {
        switch (cellFormatStr) {
            case INLINESTR:
                return new XSSFRichTextString(value).toString();
            default:
                return String.valueOf(value);
        }
    }

    private void assembleData() throws Exception {
        if (currentRowIndex >= beginReadRowIndex) {
            List<ExcelPropertyEntity> propertyList = excelMapping.getPropertyList();
            for (int i = 0; i < propertyList.size() - cellsOnRow.size(); i++) {
                cellsOnRow.add(i, "");
            }
            if (isAllEmptyRowData()) {
                throw new AllEmptyRowException("第{}行为空行,第{}个Sheet导入结束!", currentRowIndex + 1, currentSheetIndex + 1);
            }
            Object entity = excelClass.newInstance();
            ErrorEntity errorEntity = ErrorEntity.builder().build();
            for (int i = 0; i < propertyList.size(); i++) {
                dataCurrentCellIndex = i;
                Object cellValue = cellsOnRow.get(i);
                ExcelPropertyEntity property = propertyList.get(i);

                errorEntity = checkCellValue(i, property, cellValue);
                if (errorEntity.getErrorMessage() != null) {
                    break;
                }
                cellValue = convertCellValue(property, cellValue);
                if (cellValue != null) {
                    Field field = property.getFieldEntity();
                    field.set(entity, cellValue);
                }
            }
            if (errorEntity.getErrorMessage() == null) {
                importFunction.onProcess(currentSheetIndex + 1, currentRowIndex + 1, entity);
            } else {
                importFunction.onError(errorEntity);
            }
        }
    }

    private boolean isAllEmptyRowData() {
        int emptyCellCount = 0;
        for (Object cellData : cellsOnRow) {
            if (StringUtil.isBlank(cellData)) {
                emptyCellCount++;
            }
        }
        return emptyCellCount == cellsOnRow.size();
    }

    private Object convertCellValue(ExcelPropertyEntity mappingProperty, Object cellValue) throws ParseException, ExecutionException {
        Class filedClazz = mappingProperty.getFieldEntity().getType();
        if (filedClazz == Date.class) {
            if (!StringUtil.isBlank(cellValue)) {
                cellValue = parse(mappingProperty.getDateFormat(), cellValue.toString());
            } else {
                cellValue = null;
            }
        } else if (filedClazz == Short.class || filedClazz == short.class) {
            cellValue = Short.valueOf(convertNullTOZERO(cellValue));
        } else if (filedClazz == Integer.class || filedClazz == int.class) {
            cellValue = Integer.valueOf(convertNullTOZERO(cellValue));
        } else if (filedClazz == Double.class || filedClazz == double.class) {
            cellValue = Double.valueOf(convertNullTOZERO(cellValue));
        } else if (filedClazz == Long.class || filedClazz == long.class) {
            cellValue = Long.valueOf(convertNullTOZERO(cellValue));
        } else if (filedClazz == Float.class || filedClazz == float.class) {
            cellValue = Float.valueOf(convertNullTOZERO(cellValue));
        } else if (filedClazz == BigDecimal.class) {
            if (mappingProperty.getScale() == -1) {
                cellValue = new BigDecimal(convertNullTOZERO(cellValue));
            } else {
                cellValue = new BigDecimal(convertNullTOZERO(cellValue)).setScale(mappingProperty.getScale(), mappingProperty.getRoundingMode());
            }
        } else if (filedClazz != String.class) {
            throw new ExcelBootException("不支持的属性类型:{},导入失败!", filedClazz);
        }

        return cellValue;
    }


    private ErrorEntity checkCellValue(Integer cellIndex, ExcelPropertyEntity mappingProperty, Object cellValue) throws Exception {
        Boolean required = mappingProperty.getRequired();
        if (null != required && required) {
            if (null == cellValue || StringUtil.isBlank(cellValue)) {
                String validErrorMessage = String.format("第[%s]个Sheet,第[%s]行,第[%s]列必填单元格为空!"
                        , currentSheetIndex + 1, currentRowIndex + 1, cellIndex + 1);
                return buildErrorMsg(cellIndex, cellValue, validErrorMessage);
            }
        }

        String regex = mappingProperty.getRegex();
        if (!StringUtil.isBlank(cellValue) && !StringUtil.isBlank(regex)) {
            boolean matches = isMatch(regex, cellValue.toString());
            if (!matches) {
                String regularExpMessage = mappingProperty.getRegexMessage();
                String validErrorMessage = String.format("第[%s]个Sheet,第[%s]行,第[%s]列,单元格值:[%s],正则表达式[%s]校验失败!"
                        , currentSheetIndex + 1, currentRowIndex + 1, cellIndex + 1, cellValue, regularExpMessage);
                return buildErrorMsg(cellIndex, cellValue, validErrorMessage);
            }
        }

        return buildErrorMsg(cellIndex, cellValue, null);
    }

    private ErrorEntity buildErrorMsg(Integer cellIndex, Object cellValue,
                                      String validErrorMessage) {
        return ErrorEntity.builder()
                .sheetIndex(currentSheetIndex + 1)
                .rowIndex(currentRowIndex + 1)
                .cellIndex(cellIndex + 1)
                .cellValue(StringUtil.convertNull(cellValue))
                .errorMessage(validErrorMessage)
                .build();
    }

    /**
     * 计算两个单元格之间的单元格数目(同一行)
     *
     * @param refA
     * @param refB
     * @return
     */
    public int countNullCell(String refA, String refB) {
        String xfdA = refA.replaceAll("\\d+", "");
        String xfdB = refB.replaceAll("\\d+", "");

        xfdA = fillChar(xfdA, 3, '@', true);
        xfdB = fillChar(xfdB, 3, '@', true);

        char[] letterA = xfdA.toCharArray();
        char[] letterB = xfdB.toCharArray();
        int res = (letterA[0] - letterB[0]) * 26 * 26 + (letterA[1] - letterB[1]) * 26 + (letterA[2] - letterB[2]);
        return res - 1;
    }

    private String fillChar(String str, int len, char let, boolean isPre) {
        int lenA = str.length();
        if (lenA < len) {
            if (isPre) {
                StringBuilder strBuilder = new StringBuilder(str);
                for (int i = 0; i < (len - lenA); i++) {
                    strBuilder.insert(0, let);
                }
                str = strBuilder.toString();
            } else {
                StringBuilder strBuilder = new StringBuilder(str);
                for (int i = 0; i < (len - lenA); i++) {
                    strBuilder.append(let);
                }
                str = strBuilder.toString();
            }
        }
        return str;
    }


    /**
     * 单元格中的数据可能的数据类型
     */
    enum ExcelCellType {
        INLINESTR, STRING, NULL
    }


}
