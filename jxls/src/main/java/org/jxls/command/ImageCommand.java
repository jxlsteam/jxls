package org.jxls.command;

import org.jxls.area.Area;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ImageType;
import org.jxls.common.Size;
import org.jxls.transform.Transformer;

/**
 * Implements image rendering
 * Image is specified with image index in the workbook or by providing image bytes and type
 * @author Leonid Vysochyn
 */
public class ImageCommand extends AbstractCommand {

    public static final String COMMAND_NAME = "image";
    private byte[] imageBytes;
    private ImageType imageType = ImageType.PNG;
    private Area area;
    /**
     * Expression that can be evaluated to image byte array byte[]
     */
    private String src;

    public ImageCommand() {
    }

    public ImageCommand(String image, ImageType imageType) {
        this.src = image;
        this.imageType = imageType;
    }

    public ImageCommand(byte[] imageBytes, ImageType imageType) {
        this.imageBytes = imageBytes;
        this.imageType = imageType;
    }

    /**
     * @return src expression producing image byte array
     */
    public String getSrc() {
        return src;
    }

    /**
     * @param src expression resulting in image byte array
     */
    public void setSrc(String src) {
        this.src = src;
    }

    public void setImageType(String strType){
        imageType = ImageType.valueOf(strType);
    }

    @Override
    public Command addArea(Area area) {
        if( areaList.size() >= 1){
            throw new IllegalArgumentException("You can add only a single area to 'image' command");
        }
        this.area = area;
        return super.addArea(area);
    }

    public String getName() {
        return COMMAND_NAME;
    }

    public Size applyAt(CellRef cellRef, Context context) {
        if( area == null ){
            throw new IllegalArgumentException("No area is defined for image command");
        }
        Transformer transformer = getTransformer();
        Size imageAnchorAreaSize = new Size(area.getSize().getWidth() + 1, area.getSize().getHeight() + 1);
        AreaRef imageAnchorArea = new AreaRef(cellRef, imageAnchorAreaSize);
        byte[] imgBytes = imageBytes;
        if( src != null ){
            Object imgObj = getTransformationConfig().getExpressionEvaluator().evaluate(src, context.toMap());
            if( imgObj == null ){
                return area.getSize();
            }
            if( !(imgObj instanceof byte[]) ){
                throw new IllegalArgumentException("src value must contain image bytes (byte[])");
            }
            imgBytes = (byte[]) imgObj;
        }
        transformer.addImage(imageAnchorArea, imgBytes, imageType);
        return area.getSize();
    }
}
