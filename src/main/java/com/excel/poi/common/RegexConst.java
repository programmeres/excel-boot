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
public class RegexConst {
    /**
     * 金额校验:支持正负,整数位不大于12位,小数位不大于2位,也可以没有小数位
     */
    public static final String AMOUNT_REGEX = "((-|\\+?)\\d{1,12}|(-|\\+?)\\d{1,12}\\.\\d{1,2})";
    /**
     * 手机号验证:共11位,1开头
     */
    public static final String PHONE_REGEX = "^(1)\\d{10}$";
    /**
     * 邮箱验证
     */
    public static final String MAIL_REGEX = "^[\\w-]+(\\.[\\w-]+)*@[\\w-]+(\\.[\\w-]+)+$";
    /**
     * 身份证校验
     */
    public static final String IDCARD_REGEX = "(^[1-9]\\d{9}((0[1-9])|(10|11|12))(([0-2][1-9])|10|20|30|31)\\d{3}[0-9Xx]$)|" +
            "(^[1-9]\\d{9}((0[1-9])|(10|11|12))(([0-2][1-9])|10|20|30|31)\\d{3}$)";
    //^开头
    //[1-9] 第一位1-9中的一个      4
    //\\d{9} 五位数字           （省市县地区+出生年份）
    //((0[1-9])|(10|11|12))     01（月份）
    //(([0-2][1-9])|10|20|30|31)01（日期）
    //\\d{3} 三位数字            123（第十七位奇数代表男，偶数代表女）
    //[0-9Xx] 0123456789Xx其中的一个 X（第十八位为校验值）
    //$结尾

    //^开头
    //[1-9] 第一位1-9中的一个      4
    //\\d{9} 五位数字           （省市县地区+出生年份）
    //((0[1-9])|(10|11|12))     01（月份）
    //(([0-2][1-9])|10|20|30|31)01（日期）
    //\\d{3} 三位数字            123（第十五位奇数代表男，偶数代表女），15位身份证不含X
    //$结尾
}
