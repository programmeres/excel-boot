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
package com.excel.poi;

import com.excel.poi.common.Constant;
import com.excel.poi.entity.ExcelEntity;
import com.excel.poi.excel.ExcelReader;
import com.excel.poi.excel.ExcelWriter;
import com.excel.poi.exception.EasyPOIException;
import com.excel.poi.factory.ExcelMappingFactory;
import com.excel.poi.function.ExportFunction;
import com.excel.poi.function.ImportFunction;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import javax.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.xml.sax.SAXException;

/**
 * @author NingWei
 */
@Slf4j
public class EasyPoi {
    private HttpServletResponse httpServletResponse;
    private OutputStream outputStream;
    private InputStream inputStream;
    private String fileName;
    private Class excelClass;
    private Integer pageSize;
    private Integer rowAccessWindowSize;
    private Integer recordCountPerSheet;

    /**
     * 导入构造器
     *
     * @param inputStream
     * @param excelClass
     */
    protected EasyPoi(InputStream inputStream, Class excelClass) {
        this(null, null, inputStream, null, excelClass, null, null, null);
    }

    /**
     * OutputStream导出构造器,一般用于导出到ftp服务器
     *
     * @param outputStream
     * @param fileName
     * @param excelClass
     */
    protected EasyPoi(OutputStream outputStream, String fileName, Class excelClass) {
        this(null, outputStream, null, fileName, excelClass, Constant.DEFAULT_PAGE_SIZE, Constant.DEFAULT_ROW_ACCESS_WINDOW_SIZE, Constant.DEFAULT_RECORD_COUNT_PEER_SHEET);
    }

    /**
     * HttpServletResponse导出构造器,一般用于浏览器导出
     *
     * @param response
     * @param fileName
     * @param excelClass
     */
    protected EasyPoi(HttpServletResponse response, String fileName, Class excelClass) {
        this(response, null, null, fileName, excelClass, Constant.DEFAULT_PAGE_SIZE, Constant.DEFAULT_ROW_ACCESS_WINDOW_SIZE, Constant.DEFAULT_RECORD_COUNT_PEER_SHEET);
    }

    /**
     * 构造器
     *
     * @param response
     * @param outputStream
     * @param inputStream
     * @param fileName
     * @param excelClass
     * @param pageSize
     * @param rowAccessWindowSize
     * @param recordCountPerSheet
     */
    protected EasyPoi(HttpServletResponse response, OutputStream outputStream, InputStream inputStream
            , String fileName, Class excelClass, Integer pageSize, Integer rowAccessWindowSize, Integer recordCountPerSheet) {
        this.httpServletResponse = response;
        this.outputStream = outputStream;
        this.inputStream = inputStream;
        this.fileName = fileName;
        this.excelClass = excelClass;
        this.pageSize = pageSize;
        this.rowAccessWindowSize = rowAccessWindowSize;
        this.recordCountPerSheet = recordCountPerSheet;
    }

    /**
     * 通过HttpServletResponse,一般用于在浏览器中导出excel
     *
     * @param httpServletResponse
     * @param fileName
     * @param clazz
     * @return
     */
    public static EasyPoi ExportBuilder(HttpServletResponse httpServletResponse, String fileName, Class clazz) {
        return new EasyPoi(httpServletResponse, fileName, clazz);
    }

    /**
     * 通过OutputStream生成excel文件,一般用于异步导出大Excel文件到ftp服务器或本地路径
     *
     * @param outputStream
     * @param fileName
     * @param clazz
     * @return
     */
    public static EasyPoi ExportBuilder(OutputStream outputStream, String fileName, Class clazz) {
        return new EasyPoi(outputStream, fileName, clazz);
    }

    /**
     * 导入Excel文件数据
     *
     * @param inputStreamm
     * @param clazz
     * @return
     */
    public static EasyPoi ImportBuilder(InputStream inputStreamm, Class clazz) {
        return new EasyPoi(inputStreamm, clazz);
    }

    /**
     * 用于浏览器导出
     *
     * @param param
     * @param exportFunction
     * @param ExportFunction
     * @param <R>
     * @param <T>
     */
    public <R, T> void exportResponse(R param, ExportFunction<R, T> exportFunction) {
        try {
            if (httpServletResponse == null) {
                throw new EasyPOIException("httpServletResponse参数为空!");
            }
            VerifyParams();
            ExcelEntity excelEntity = ExcelMappingFactory.loadExportExcelClass(excelClass);
            excelEntity.setFileName(fileName);
            ExcelWriter excelWriter = new ExcelWriter(excelEntity, pageSize, rowAccessWindowSize, recordCountPerSheet);
            SXSSFWorkbook workbook = excelWriter.generateWorkbook(param, exportFunction);
            download(workbook, httpServletResponse, URLEncoder.encode(fileName + ".xlsx", "UTF-8"));
        } catch (Exception e) {
            throw new EasyPOIException(e);
        }
    }

    /**
     * 用于浏览器分sheet导出
     *
     * @param param
     * @param exportFunction
     * @param ExportFunction
     * @param <R>
     * @param <T>
     */
    public <R, T> void exportMultiSheetResponse(R param, ExportFunction<R, T> exportFunction) {
        try {
            if (httpServletResponse == null) {
                throw new EasyPOIException("httpServletResponse参数为空!");
            }
            VerifyParams();
            ExcelEntity excelMapping = ExcelMappingFactory.loadExportExcelClass(excelClass);
            excelMapping.setFileName(fileName);
            ExcelWriter excelWriter = new ExcelWriter(excelMapping, pageSize, rowAccessWindowSize, recordCountPerSheet);
            SXSSFWorkbook workbook = excelWriter.generateMultiSheetWorkbook(param, exportFunction);
            download(workbook, httpServletResponse, URLEncoder.encode(fileName + ".xlsx", "UTF-8"));
        } catch (Exception e) {
            throw new EasyPOIException(e);
        }
    }

    /**
     * 通过OutputStream导出excel文件,一般用于异步导出大Excel文件到本地路径
     *
     * @param param
     * @param ExportFunction
     * @param exportFunction
     * @param <R>
     * @param <T>
     */
    public <R, T> void exportStream(R param, ExportFunction<R, T> exportFunction) {
        try {
            write(generateExcelStream(param, exportFunction));
        } catch (Exception e) {
            throw new EasyPOIException(e);
        }
    }

    /**
     * 通过OutputStream导出excel文件,一般用于异步导出大Excel文件到ftp服务器
     *
     * @param param
     * @param exportFunction
     * @param ExportFunction
     * @param <R>
     * @param <T>
     * @return
     */
    public <R, T> OutputStream generateExcelStream(R param, ExportFunction<R, T> exportFunction) {
        try {
            if (outputStream == null) {
                throw new EasyPOIException("outputStream参数为空!");
            }
            VerifyParams();
            ExcelEntity excelMapping = ExcelMappingFactory.loadExportExcelClass(excelClass);
            excelMapping.setFileName(fileName);
            ExcelWriter excelWriter = new ExcelWriter(excelMapping, pageSize, rowAccessWindowSize, recordCountPerSheet);
            SXSSFWorkbook workbook = excelWriter.generateWorkbook(param, exportFunction);
            workbook.write(outputStream);
            return outputStream;
        } catch (Exception e) {
            throw new EasyPOIException(e);
        }
    }

    /**
     * 通过OutputStream分sheet导出excel文件,一般用于异步导出大Excel文件到本地路径
     *
     * @param param
     * @param ExportFunction
     * @param exportFunction
     * @param <R>
     * @param <T>
     */
    public <R, T> void exportMultiSheetStream(R param, ExportFunction<R, T> exportFunction) {
        try {
            write(generateMultiSheetExcelStream(param, exportFunction));
        } catch (Exception e) {
            throw new EasyPOIException(e);
        }
    }


    /**
     * 通过OutputStream分sheet导出excel文件,一般用于异步导出大Excel文件到ftp服务器
     *
     * @param param
     * @param exportFunction
     * @param ExportFunction
     * @param <R>
     * @param <T>
     * @return
     */
    public <R, T> OutputStream generateMultiSheetExcelStream(R param, ExportFunction<R, T> exportFunction) {
        try {
            if (outputStream == null) {
                throw new EasyPOIException("outputStream参数为空!");
            }
            VerifyParams();
            ExcelEntity excelMapping = ExcelMappingFactory.loadExportExcelClass(excelClass);
            excelMapping.setFileName(fileName);
            ExcelWriter excelWriter = new ExcelWriter(excelMapping, pageSize, rowAccessWindowSize, recordCountPerSheet);
            SXSSFWorkbook workbook = excelWriter.generateMultiSheetWorkbook(param, exportFunction);
            workbook.write(outputStream);
            return outputStream;
        } catch (Exception e) {
            throw new EasyPOIException(e);
        }
    }

    /**
     * 导出-导入模板
     *
     * @param data
     * @throws Exception
     */
    public void exportTemplate() {
        try {
            if (httpServletResponse == null) {
                throw new EasyPOIException("httpServletResponse参数为空!");
            }
            VerifyParams();
            ExcelEntity excelMapping = ExcelMappingFactory.loadExportExcelClass(excelClass);
            excelMapping.setFileName(fileName);
            ExcelWriter excelWriter = new ExcelWriter(excelMapping, pageSize, rowAccessWindowSize, recordCountPerSheet);
            SXSSFWorkbook workbook = excelWriter.generateTemplateWorkbook();
            download(workbook, httpServletResponse, URLEncoder.encode(fileName + ".xlsx", "UTF-8"));
        } catch (Exception e) {
            throw new EasyPOIException(e);
        }
    }

    /**
     * 导入excel全部sheet
     *
     * @param inputStream
     * @param importFunction
     * @throws OpenXML4JException
     * @throws SAXException
     * @throws IOException
     */
    public void importExcel(ImportFunction importFunction) {
        try {
            if (importFunction == null) {
                throw new EasyPOIException("excelReadHandler参数为空!");
            }
            if (inputStream == null) {
                throw new EasyPOIException("inputStream参数为空!");
            }

            ExcelEntity excelMapping = ExcelMappingFactory.loadImportExcelClass(excelClass);
            ExcelReader excelReader = new ExcelReader(excelClass, excelMapping, importFunction);
            excelReader.process(inputStream);
        } catch (Exception e) {
            throw new EasyPOIException(e);
        }

    }

    /**
     * 生成文件
     *
     * @param out
     * @throws IOException
     */
    private void write(OutputStream out) throws IOException {
        if (null != out) {
            try {
                out.flush();
            } finally {
                out.close();
            }
        }
    }

    /**
     * 生成文件
     *
     * @param wb
     * @param out
     * @throws IOException
     */
    private void write(SXSSFWorkbook wb, OutputStream out) throws IOException {
        if (null != out) {
            try {
                wb.write(out);
                out.flush();
            } finally {
                out.close();
            }
        }
    }


    /**
     * 构建Excel服务器响应格式
     *
     * @param wb
     * @param response
     * @param filename
     * @throws IOException
     */
    private void download(SXSSFWorkbook wb, HttpServletResponse response, String filename) throws IOException {
        OutputStream out = response.getOutputStream();
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-disposition",
                String.format("attachment; filename=%s", filename));
        write(wb, out);
    }

    private void VerifyParams() {
        if (excelClass == null) {
            throw new EasyPOIException("excelClass参数为空!");
        }
        if (fileName == null) {
            throw new EasyPOIException("fileName参数为空!");
        }
    }

}