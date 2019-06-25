package com.excel.poi.common;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.UnsupportedEncodingException;

/**根据RFC 5987规范生成disposition值, 解决浏览器兼容以及中文乱码问题
 * @author BiLuohen
 * @date 6/25/2019
 */
public class WebFilenameUtils {

    private static final Logger LOGGER = LoggerFactory.getLogger(WebFilenameUtils.class);

    private static final String DISPOSITION_FORMAT = "attachment; filename=\"%s\"; filename*=utf-8''%s";

    /**
     * 未编码文件名转Content-Disposition值
     *
     * @param filename 未编码的文件名(包含文件后缀)
     * @return Content-Disposition值
     */
    public static String disposition(String filename) {
        String codedFilename = filename;
        try {
            if (!StringUtil.isBlank(filename)) {
                codedFilename = java.net.URLEncoder.encode(filename, "UTF-8");
            }
        } catch (UnsupportedEncodingException e) {
            LOGGER.error("不支持的编码:", e);
        }
        return String.format(DISPOSITION_FORMAT, codedFilename, codedFilename);

    }
}
