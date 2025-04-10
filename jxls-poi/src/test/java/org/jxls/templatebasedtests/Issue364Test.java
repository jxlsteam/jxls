package org.jxls.templatebasedtests;

import org.apache.commons.io.IOUtils;
import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author qiuxs
 * @description
 * @date 2025/4/10
 */
public class Issue364Test {

    @Test
    public void test() throws IOException {
        final InputStream imageStream = Issue364Test.class.getResourceAsStream("/org/jxls/examples/car.jpg");
        final byte[] imageBytes = IOUtils.toByteArray(imageStream);
        Context context = new Context();
        context.putVar("image", imageBytes);

        JxlsHelper.getInstance().processTemplate(Issue364Test.class.getResourceAsStream("/org/jxls/templatebasedtests/issue364Test.xlsx"), new FileOutputStream("./issueTestOutput.xlsx"), context);
    }

}
