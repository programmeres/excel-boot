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
package com.excel.poi.common;

/**
 * @author NingWei
 */
public class StringUtil {
    /**
     * 判断字符串是否为空
     *
     * @param str
     * @return
     */
    public static boolean isBlank(Object str) {
        return str == null || "".equals(str.toString().trim()) || "null".equalsIgnoreCase(str.toString().trim());
    }

    /**
     * 格式化null为空
     *
     * @param str
     * @return
     */
    public static String convertNull(Object str) {
        if (str == null) {
            return "";
        }
        return str.toString();
    }

    /**
     * 格式化null为0
     *
     * @param str
     * @return
     */
    public static String convertNullTOZERO(Object str) {
        if (str == null || "".equals(str.toString().trim()) || "null".equalsIgnoreCase(str.toString().trim())) {
            return "0";
        }
        return str.toString();
    }

    public static boolean isNumeric(CharSequence cs) {
        if (isBlank(cs)) {
            return false;
        } else {
            int sz = cs.length();

            for (int i = 0; i < sz; ++i) {
                if (!Character.isDigit(cs.charAt(i))) {
                    return false;
                }
            }

            return true;
        }
    }
}
