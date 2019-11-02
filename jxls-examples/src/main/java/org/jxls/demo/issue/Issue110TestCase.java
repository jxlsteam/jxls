package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Conditional formatting copying issue
 */
public class Issue110TestCase {
    
    private final static String INPUT_FILE_PATH = "issue110_template.xlsx";
    private final static String OUTPUT_FILE_PATH = "target/issue110_output.xlsx";

    public static void main(String[] args) throws IOException {
        try(InputStream is = Issue110TestCase.class.getResourceAsStream(INPUT_FILE_PATH)) {
            try (OutputStream os = new FileOutputStream(OUTPUT_FILE_PATH)) {
                List<Item> items = new ArrayList<>();
                items.add(new Item("X", 1, 2));
                items.add(new Item("Y", 3, 4));
                items.add(new Item("Z", 5, 6));
                Context context = new Context();
                context.putVar("items", items);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

    public static class Item {
        public final String label;
        public final int a;
        public final int b;

        public Item(String label, int a, int b) {
            this.label = label;
            this.a = a;
            this.b = b;
        }
    }
}
