/*
 * Copyright 2018 Badari Narayana
 *
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
 */


package com.kbn.excel;

import org.robotframework.javalib.library.AnnotationLibrary;
import org.robotframework.javalib.library.KeywordDocumentationRepository;
import org.robotframework.javalib.library.RobotJavaLibrary;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;

public class JyExcelLibrary implements KeywordDocumentationRepository, RobotJavaLibrary{
    public static final String ROBOT_LIBRARY_SCOPE = "GLOBAL";
    public static JyExcelLibrary instance;
    private final AnnotationLibrary annotationLibrary = new AnnotationLibrary(
            "com/kbn/excel/keyword/**.class");
    private static final String LIBRARY_DOCUMENTATION = "JyExcelLibrary is a Robot Framework test library for handling excel sheets using Jython.\n"+
            "Currently it supports .xlsx format only. This cannot be used with python.\n";

    public JyExcelLibrary() {
        this(Collections.<String> emptyList());
    }

    protected JyExcelLibrary(final String keywordPattern) {
        this(new ArrayList<String>() {
            {
                add(keywordPattern);
            }
        });
    }

    protected JyExcelLibrary(Collection<String> keywordPatterns) {
        addKeywordPatterns(keywordPatterns);
        instance = this;
    }

    private void addKeywordPatterns(Collection<String> keywordPatterns) {
        for (String pattern : keywordPatterns) {
            annotationLibrary.addKeywordPattern(pattern);
        }
    }


    public Object runKeyword(String keywordName, Object[] args) {
        return annotationLibrary.runKeyword(keywordName, toStrings(args));
    }


    public String[] getKeywordArguments(String keywordName) {
        return annotationLibrary.getKeywordArguments(keywordName);
    }


    public String getKeywordDocumentation(String keywordName) {
        if (keywordName.equals("__intro__"))
            return LIBRARY_DOCUMENTATION;
        return annotationLibrary.getKeywordDocumentation(keywordName);
    }


    public String[] getKeywordNames() {
        return annotationLibrary.getKeywordNames();
    }


    private Object[] toStrings(Object[] args) {
        Object[] newArgs = new Object[args.length];
        for (int i = 0; i < newArgs.length; i++) {
            if (args[i].getClass().isArray()) {
                newArgs[i] = args[i];
            } else {
                newArgs[i] = args[i].toString();
            }
        }
        return newArgs;
    }

}
