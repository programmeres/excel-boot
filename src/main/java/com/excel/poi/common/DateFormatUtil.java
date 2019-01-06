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

import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.ExecutionException;

/**
 * @author NingWei
 */
public class DateFormatUtil {

    private final static LoadingCache<String, SimpleDateFormat> loadCache =
            CacheBuilder.newBuilder()
                    .maximumSize(5)
                    .build(new CacheLoader<String, SimpleDateFormat>() {
                        @Override
                        public SimpleDateFormat load(String pattern) {
                            return new SimpleDateFormat(pattern);
                        }
                    });

    public static Date parse(String pattern, String value) throws ExecutionException, ParseException {
        return loadCache.get(pattern).parse(value);
    }

    public static String format(String pattern, Date value) throws ExecutionException {
        return loadCache.get(pattern).format(value);
    }
}
