package org.jxls.templatebasedtests;

import org.apache.commons.io.IOUtils;
import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
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
        final InputStream imageStream = Issue364Test.class.getResourceAsStream("/org/jxls/examples/tulip.jpg");
        final BufferedImage image = ImageIO.read(imageStream);
        final BufferedImage resizeImage = this.resizeWithAspectRatio(image, 400, 400);
        final ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(resizeImage, "jpg", baos);
        final byte[] imageBytes = baos.toByteArray();
        Context context = new Context();
        context.putVar("image", imageBytes);

        JxlsHelper.getInstance().processTemplate(Issue364Test.class.getResourceAsStream("/org/jxls/templatebasedtests/issue364Test.xlsx"), new FileOutputStream("./issueTestOutput.xlsx"), context);
    }

    private BufferedImage resizeWithAspectRatio(BufferedImage image, int maxWidth, int maxHeight) throws IOException {
        int originalWidth = image.getWidth();
        int originalHeight = image.getHeight();
        // 本来就小的没必要调整大小
        if (originalWidth < maxWidth && originalHeight < maxHeight) {
            return image;
        }

        // 计算保持宽高比的尺寸
        float aspectRatio = (float) originalWidth / originalHeight;
        int newWidth = maxWidth;
        int newHeight = (int) (maxWidth / aspectRatio);

        if (newHeight > maxHeight) {
            newHeight = maxHeight;
            newWidth = (int) (maxHeight * aspectRatio);
        }

        BufferedImage resizedImage = new BufferedImage(
                newWidth, newHeight, BufferedImage.TYPE_INT_RGB);

        Graphics2D g = resizedImage.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_INTERPOLATION,
                RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g.drawImage(image, 0, 0, newWidth, newHeight, null);
        g.dispose();

        return resizedImage;
    }

}
