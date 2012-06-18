package com.jxls.writer.command;

import com.jxls.writer.area.Area;
import com.jxls.writer.common.*;
import com.jxls.writer.transform.Transformer;

/**
 * @author Leonid Vysochyn
 */
public class ImageCommand extends AbstractCommand {

    byte[] imageBytes;
    ImageType imageType;
    Area area;
    Integer imageIdx;
    String imgBean;

    public ImageCommand() {
    }

    public ImageCommand(String imgBean, ImageType imageType) {
        this.imgBean = imgBean;
        this.imageType = imageType;
    }

    public ImageCommand(Integer imageIdx) {
        this.imageIdx = imageIdx;
    }

    public ImageCommand(byte[] imageBytes, ImageType imageType) {
        this.imageBytes = imageBytes;
        this.imageType = imageType;
    }

    public String getImgBean() {
        return imgBean;
    }

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
        if( imageIdx != null ){
            transformer.addImage(areaRef, imageIdx);
        }else{
            byte[] imgBytes = imageBytes;
            if( imgBean != null ){
                Object imgObj = context.getVar(imgBean);
                if( !(imgObj instanceof byte[]) ){
                    throw new IllegalArgumentException("imgBean value must contain image bytes (byte[])");
                }
                imgBytes = (byte[]) imgObj;
            }
            transformer.addImage(areaRef, imgBytes, imageType);
        }
        return area.getSize();
    }
}
