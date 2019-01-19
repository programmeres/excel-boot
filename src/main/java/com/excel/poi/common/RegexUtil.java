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
import java.util.concurrent.ExecutionException;
import java.util.regex.Pattern;

/**
 * @author NingWei
 */
public class RegexUtil {

    private final static LoadingCache<String, Pattern> LOAD_CACHE =
            CacheBuilder.newBuilder()
                    .maximumSize(30)
                    .build(new CacheLoader<String, Pattern>() {
                        @Override
                        public Pattern load(String pattern) {
                            return Pattern.compile(pattern);
                        }
                    });

    public static Boolean isMatch(String pattern, String value) throws ExecutionException {
        return LOAD_CACHE.get(pattern).matcher(value).matches();
    }
}
