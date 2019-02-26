package org.jxls.demo;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ShiftWithEmptyList {

    final static List<String> NON_EMPTY_LIST = new ArrayList<String>(Arrays.asList("A", "B", "C"));
    final static List<String> EMPTY_LIST = new ArrayList<String>();

    final static String EMPTY_LIST_NAME = "emptyList";
    final static String NON_EMPTY_LIST_NAME = "nonEmptyList";

    final static String INPUT_FILE_PATH = "emptylist_shift.xls";
    final static String OUTPUT_FILE_PATH = "target/emptylist_shift_output.xls";

    public static void main(String[] args) throws IOException {
        try(InputStream is = ShiftStrategyDemo.class.getResourceAsStream(INPUT_FILE_PATH)) {
            try (OutputStream os = new FileOutputStream(OUTPUT_FILE_PATH)) {
                Context context = new Context();
                context.putVar(NON_EMPTY_LIST_NAME, NON_EMPTY_LIST);
                context.putVar(EMPTY_LIST_NAME, EMPTY_LIST);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

}