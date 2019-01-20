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
import com.excel.poi.exception.ExcelBootException;
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
public class ExcelBoot {
    private HttpServletResponse httpServletResponse;
    private OutputStream outputStream;
    private InputStream inputStream;
    private String fileName;
    private Class excelClass;
    private Integer pageSize;
    private Integer rowAccessWindowSize;
    private Integer recordCountPerSheet;
    private Boolean openAutoColumWidth;

    /**
     * 导入构造器
     *
     * @param inputStream
     * @param excelClass
     */
    protected ExcelBoot(InputStream inputStream, Class excelClass) {
        this(null, null, inputStream, null, excelClass, null, null, null, null);
    }

    /**
     * OutputStream导出构造器,一般用于导出到ftp服务器
     *
     * @param outputStream
     * @param fileName
     * @param excelClass
     */
    protected ExcelBoot(OutputStream outputStream, String fileName, Class excelClass) {
        this(null, outputStream, null, fileName, excelClass, Constant.DEFAULT_PAGE_SIZE, Constant.DEFAULT_ROW_ACCESS_WINDOW_SIZE, Constant.DEFAULT_RECORD_COUNT_PEER_SHEET, Constant.OPEN_AUTO_COLUM_WIDTH);
    }

    /**
     * HttpServletResponse导出构造器,一般用于浏览器导出
     *
     * @param response
     * @param fileName
     * @param excelClass
     */
    protected ExcelBoot(HttpServletResponse response, String fileName, Class excelClass) {
        this(response, null, null, fileName, excelClass, Constant.DEFAULT_PAGE_SIZE, Constant.DEFAULT_ROW_ACCESS_WINDOW_SIZE, Constant.DEFAULT_RECORD_COUNT_PEER_SHEET, Constant.OPEN_AUTO_COLUM_WIDTH);
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
     * @param openAutoColumWidth
     */
    protected ExcelBoot(HttpServletResponse response, OutputStream outputStream, InputStream inputStream
            , String fileName, Class excelClass, Integer pageSize, Integer rowAccessWindowSize, Integer recordCountPerSheet, Boolean openAutoColumWidth) {
        this.httpServletResponse = response;
        this.outputStream = outputStream;
        this.inputStream = inputStream;
        this.fileName = fileName;
        this.excelClass = excelClass;
        this.pageSize = pageSize;
        this.rowAccessWindowSize = rowAccessWindowSize;
        this.recordCountPerSheet = recordCountPerSheet;
        this.openAutoColumWidth = openAutoColumWidth;
    }

    /**
     * 通过HttpServletResponse,一般用于在浏览器中导出excel
     *
     * @param httpServletResponse
     * @param fileName
     * @param clazz
     * @return
     */
    public static ExcelBoot ExportBuilder(HttpServletResponse httpServletResponse, String fileName, Class clazz) {
        return new ExcelBoot(httpServletResponse, fileName, clazz);
    }

    /**
     * 通过OutputStream生成excel文件,一般用于异步导出大Excel文件到ftp服务器或本地路径
     *
     * @param outputStream
     * @param fileName
     * @param clazz
     * @return
     */
    public static ExcelBoot ExportBuilder(OutputStream outputStream, String fileName, Class clazz) {
        return new ExcelBoot(outputStream, fileName, clazz);
    }

    /**
     * HttpServletResponse 通用导出Excel构造器
     *
     * @param response
     * @param fileName
     * @param excelClass
     * @param pageSize
     * @param rowAccessWindowSize
     * @param recordCountPerSheet
     * @param openAutoColumWidth
     * @return
     */
    public static ExcelBoot ExportBuilder(HttpServletResponse response, String fileName, Class excelClass,
                                          Integer pageSize, Integer rowAccessWindowSize, Integer recordCountPerSheet, Boolean openAutoColumWidth) {
        return new ExcelBoot(response, null, null
                , fileName, excelClass, pageSize, rowAccessWindowSize, recordCountPerSheet, openAutoColumWidth);
    }

    /**
     * OutputStream 通用导出Excel构造器
     *
     * @param outputStream
     * @param fileName
     * @param excelClass
     * @param pageSize
     * @param rowAccessWindowSize
     * @param recordCountPerSheet
     * @param openAutoColumWidth
     * @return
     */
    public static ExcelBoot ExportBuilder(OutputStream outputStream, String fileName, Class excelClass, Integer pageSize
            , Integer rowAccessWindowSize, Integer recordCountPerSheet, Boolean openAutoColumWidth) {
        return new ExcelBoot(null, outputStream, null
                , fileName, excelClass, pageSize, rowAccessWindowSize, recordCountPerSheet, openAutoColumWidth);
    }

    /**
     * 导入Excel文件数据
     *
     * @param inputStreamm
     * @param clazz
     * @return
     */
    public static ExcelBoot ImportBuilder(InputStream inputStreamm, Class clazz) {
        return new ExcelBoot(inputStreamm, clazz);
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
        SXSSFWorkbook sxssfWorkbook = null;
        try {
            try {
                verifyResponse();
                sxssfWorkbook = commonSingleSheet(param, exportFunction);
                download(sxssfWorkbook, httpServletResponse, URLEncoder.encode(fileName + ".xlsx", "UTF-8"));
            } finally {
                if (sxssfWorkbook != null) {
                    sxssfWorkbook.close();
                }
                if (httpServletResponse.getOutputStream() != null) {
                    httpServletResponse.getOutputStream().close();
                }
            }
        } catch (Exception e) {
            throw new ExcelBootException(e);
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
        OutputStream outputStream = null;
        try {
            try {
                outputStream = generateStream(param, exportFunction);
                write(outputStream);
            } finally {
                if (outputStream != null) {
                    outputStream.close();
                }
            }
        } catch (Exception e) {
            throw new ExcelBootException(e);
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
    public <R, T> OutputStream generateStream(R param, ExportFunction<R, T> exportFunction) throws IOException {
        SXSSFWorkbook sxssfWorkbook = null;
        try {
            verifyStream();
            sxssfWorkbook = commonSingleSheet(param, exportFunction);
            sxssfWorkbook.write(outputStream);
            return outputStream;
        } catch (Exception e) {
            log.error("生成Excel发生异常! 异常信息:", e);
            if (sxssfWorkbook != null) {
                sxssfWorkbook.close();
            }
            throw new ExcelBootException(e);
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
        SXSSFWorkbook sxssfWorkbook = null;
        try {
            try {
                verifyResponse();
                sxssfWorkbook = commonMultiSheet(param, exportFunction);
                download(sxssfWorkbook, httpServletResponse, URLEncoder.encode(fileName + ".xlsx", "UTF-8"));
            } finally {
                if (sxssfWorkbook != null) {
                    sxssfWorkbook.close();
                }
            }
        } catch (Exception e) {
            throw new ExcelBootException(e);
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
        OutputStream outputStream = null;
        try {
            try {
                outputStream = generateMultiSheetStream(param, exportFunction);
                write(outputStream);
            } finally {
                if (outputStream != null) {
                    outputStream.close();
                }
            }
        } catch (Exception e) {
            throw new ExcelBootException(e);
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
    public <R, T> OutputStream generateMultiSheetStream(R param, ExportFunction<R, T> exportFunction) throws IOException {
        SXSSFWorkbook sxssfWorkbook = null;
        try {
            verifyStream();
            sxssfWorkbook = commonMultiSheet(param, exportFunction);
            sxssfWorkbook.write(outputStream);
            return outputStream;
        } catch (Exception e) {
            log.error("分Sheet生成Excel发生异常! 异常信息:", e);
            if (sxssfWorkbook != null) {
                sxssfWorkbook.close();
            }
            throw new ExcelBootException(e);
        }
    }

    /**
     * 导出-导入模板
     *
     * @param data
     * @throws Exception
     */
    public void exportTemplate() {
        SXSSFWorkbook sxssfWorkbook = null;
        try {
            try {
                verifyResponse();
                verifyParams();
                ExcelEntity excelMapping = ExcelMappingFactory.loadExportExcelClass(excelClass, fileName);
                ExcelWriter excelWriter = new ExcelWriter(excelMapping, pageSize, rowAccessWindowSize, recordCountPerSheet, openAutoColumWidth);
                sxssfWorkbook = excelWriter.generateTemplateWorkbook();
                download(sxssfWorkbook, httpServletResponse, URLEncoder.encode(fileName + ".xlsx", "UTF-8"));
            } finally {
                if (sxssfWorkbook != null) {
                    sxssfWorkbook.close();
                }
                if (httpServletResponse.getOutputStream() != null) {
                    httpServletResponse.getOutputStream().close();
                }
            }
        } catch (Exception e) {
            throw new ExcelBootException(e);
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
                throw new ExcelBootException("excelReadHandler参数为空!");
            }
            if (inputStream == null) {
                throw new ExcelBootException("inputStream参数为空!");
            }

            ExcelEntity excelMapping = ExcelMappingFactory.loadImportExcelClass(excelClass);
            ExcelReader excelReader = new ExcelReader(excelClass, excelMapping, importFunction);
            excelReader.process(inputStream);
        } catch (Exception e) {
            throw new ExcelBootException(e);
        }

    }

    private <R, T> SXSSFWorkbook commonSingleSheet(R param, ExportFunction<R, T> exportFunction) throws Exception {
        verifyParams();
        ExcelEntity excelMapping = ExcelMappingFactory.loadExportExcelClass(excelClass, fileName);
        ExcelWriter excelWriter = new ExcelWriter(excelMapping, pageSize, rowAccessWindowSize, recordCountPerSheet, openAutoColumWidth);
        return excelWriter.generateWorkbook(param, exportFunction);
    }

    private <R, T> SXSSFWorkbook commonMultiSheet(R param, ExportFunction<R, T> exportFunction) throws Exception {
        verifyParams();
        ExcelEntity excelMapping = ExcelMappingFactory.loadExportExcelClass(excelClass, fileName);
        ExcelWriter excelWriter = new ExcelWriter(excelMapping, pageSize, rowAccessWindowSize, recordCountPerSheet, openAutoColumWidth);
        return excelWriter.generateMultiSheetWorkbook(param, exportFunction);
    }

    /**
     * 生成文件
     *
     * @param out
     * @throws IOException
     */
    private void write(OutputStream out) throws IOException {
        if (null != out) {
            out.flush();
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
        if (null != out) {
            wb.write(out);
            out.flush();
        }
    }

    private void verifyResponse() {
        if (httpServletResponse == null) {
            throw new ExcelBootException("httpServletResponse参数为空!");
        }
    }

    private void verifyStream() {
        if (outputStream == null) {
            throw new ExcelBootException("outputStream参数为空!");
        }
    }

    private void verifyParams() {
        if (excelClass == null) {
            throw new ExcelBootException("excelClass参数为空!");
        }
        if (fileName == null) {
            throw new ExcelBootException("fileName参数为空!");
        }
    }

}