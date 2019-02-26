package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Issue103TestCase {
// --------- SETTINGS ---------

    // define the lists which would be used in the template
    final static List<String> NON_EMPTY_LIST = new ArrayList<String>(Arrays.asList("A", "B", "C", "D", "E", "F", "G", "H", "I", "J"));
    final static List<String> EMPTY_LIST = new ArrayList<String>();

    // define how the lists would be named in the template
    final static String EMPTY_LIST_NAME = "emptyList";
    final static String NON_EMPTY_LIST_NAME = "nonEmptyList";

    final static String INPUT_FILE_PATH = "Issue103_Template.xls";
    final static String OUTPUT_FILE_PATH = "target/Issue103_Output.xls";

    // --------- -------- ---------

    public static void main(String[] args) throws IOException {
        try(InputStream is = Issue103TestCase.class.getResourceAsStream(INPUT_FILE_PATH)) {
            try (OutputStream os = new FileOutputStream(OUTPUT_FILE_PATH)) {
                Context context = new Context();
                context.putVar(NON_EMPTY_LIST_NAME, NON_EMPTY_LIST);
                context.putVar(EMPTY_LIST_NAME, EMPTY_LIST);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }

    }

}
