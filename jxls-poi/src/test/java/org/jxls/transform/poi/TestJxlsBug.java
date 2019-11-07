package org.jxls.transform.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

public class TestJxlsBug {
    
    public static void main(String[] args) throws IOException {
        URL xlsUrl = TestJxlsBug.class.getResource("test_bug.xls");
        String filePath = xlsUrl.getFile();
        int i = filePath.lastIndexOf('/');
        String targetFilePath = filePath.substring(0, i) + "/target.xls";
        List<Integer> data = new ArrayList<>();
        for (int j = 1; j <= 3; j++) {
            data.add(j);
        }
        Context context = new Context();
        context.putVar("data", data);
        FileOutputStream os = new FileOutputStream(targetFilePath);
        InputStream is = xlsUrl.openStream();
        JxlsHelper.getInstance().processTemplate(is, os, context);
        is.close();
        os.close();
        System.out.println("Finish");
    }
}
