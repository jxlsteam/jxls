package org.jxls.demo.issue;

import org.apache.commons.io.IOUtils;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * A modification of fangzhengjin example
 */
public class Issue159TestCase {

    public static void main(String[] args) throws IOException {
        try (InputStream template = Issue159TestCase.class.getResourceAsStream("Issue159_template.xlsx")) {
            try (InputStream stamp = Issue159TestCase.class.getResourceAsStream("stamp.png")) {
                FileOutputStream fileOutputStream = new FileOutputStream("issue159_output.xlsx");
                Map<String, Object> model = new HashMap<>();
                List<Integer> details = new ArrayList<>();
                for (int i = 0; i < 10; i++) {
                    details.add(i);
                }
                model.put("details", details);
                assert stamp != null;
                model.put("stampImage", IOUtils.toByteArray(stamp));
                model.put("name", "name111111111");
                model.put("remark", "remark remark remark remark remark remark\n remark remark remark remark remark remark\n remark remark remark remark remark remark remark remark remark remark ");
                exportExcel(template, fileOutputStream, model);
                fileOutputStream.flush();
                fileOutputStream.close();
            }
        }
    }

    public static void exportExcel(InputStream inputStream, OutputStream outputStream, Map<String, Object> model) throws IOException {
        try {
            Context context = PoiTransformer.createInitialContext();
            if (model != null) {
                model.keySet().forEach(x -> context.putVar(x, model.get(x)));
            }
            JxlsHelper jxlsHelper = JxlsHelper.getInstance();
            XlsCommentAreaBuilder builder = new XlsCommentAreaBuilder();

            Transformer transformer = jxlsHelper.createTransformer(inputStream, outputStream);
            jxlsHelper.setUseFastFormulaProcessor(false).processTemplate(context, transformer);
        } finally {
            inputStream.close();
        }
    }
}
