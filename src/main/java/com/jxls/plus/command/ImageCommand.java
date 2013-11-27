package com.jxls.plus.command;

import com.jxls.plus.area.Area;
import com.jxls.plus.common.*;
import com.jxls.plus.transform.Transformer;

/**
 * Implements image rendering
 * Image is specified with image index in the workbook or by providing image bytes and type
 * @author Leonid Vysochyn
 */
public class ImageCommand extends AbstractCommand {

    byte[] imageBytes;
    ImageType imageType;
    Area area;
    /**
     * Image bean name in the context
     */
    String imgBean;

    public ImageCommand() {
    }

    public ImageCommand(String imgBean, ImageType imageType) {
        this.imgBean = imgBean;
        this.imageType = imageType;
    }

    public ImageCommand(byte[] imageBytes, ImageType imageType) {
        this.imageBytes = imageBytes;
        this.imageType = imageType;
    }

    /**
     * @return image bean name in the context
     */
    public String getImgBean() {
        return imgBean;
    }

    /**
     * @param imgBean image bean name in the context
     */
    public void setImgBean(String imgBean) {
        this.imgBean = imgBean;
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
        return "image";
    }

    public Size applyAt(CellRef cellRef, Context context) {
        if( area == null ){
            throw new IllegalArgumentException("No area is defined for image command");
        }
        Transformer transformer = getTransformer();
        AreaRef areaRef = new AreaRef(cellRef, area.getSize());
        byte[] imgBytes = imageBytes;
        if( imgBean != null ){
            Object imgObj = context.getVar(imgBean);
            if( !(imgObj instanceof byte[]) ){
                throw new IllegalArgumentException("imgBean value must contain image bytes (byte[])");
            }
            imgBytes = (byte[]) imgObj;
        }
        transformer.addImage(areaRef, imgBytes, imageType);
        return area.getSize();
    }
}
