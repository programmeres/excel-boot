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
public final class Constant {
    //Excel自动刷新到磁盘的数量
    public static final int DEFAULT_ROW_ACCESS_WINDOW_SIZE = 2000;
    //分页条数
    public static final int DEFAULT_PAGE_SIZE = 3000;
    //分Sheet条数
    public static final int DEFAULT_RECORD_COUNT_PEER_SHEET = 80000;

    public static final String CELL = "c";
    public static final String XYZ_LOCATION = "r";
    public static final String CELL_T_PROPERTY = "t";
    public static final String CELL_S_VALUE = "s";
    public static final String ROW = "row";

    public static final int CHINESES_ATUO_SIZE_COLUMN_WIDTH_MAX = 60;
    public static final int CHINESES_ATUO_SIZE_COLUMN_WIDTH_MIN = 15;

    public static final int MAX_RECORD_COUNT_PEER_SHEET = 1000000;
}
