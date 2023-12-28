package org.jxls.command;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jxls.area.Area;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ImageType;
import org.jxls.common.Size;
import org.jxls.transform.poi.PoiTransformer;

/**
 * <p>Implements image rendering</p>
 * <p>Image is specified by providing image bytes and type.</p>
 * 
 * @author Leonid Vysochyn
 */
public class ImageCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "image";
    private byte[] imageBytes;
    private ImageType imageType = ImageType.PNG;
    private Area area;
    /** Expression that can be evaluated to image byte array byte[] */
    private String src;
    /**
     * org.apache.poi.ss.usermodel.Picture#resize(double scaleX, double scaleY)
     * <p>
     * Resize the image.
     * <p>
     * Please note, that this method works correctly only for workbooks
     * with the default font size (Arial 10pt for .xls and Calibri 11pt for .xlsx).
     * If the default font is changed the resized image can be streched vertically or horizontally.
     * </p>
     * <p>
     * <code>resize(1.0,1.0)</code> keeps the original size,<br>
     * <code>resize(0.5,0.5)</code> resize to 50% of the original,<br>
     * <code>resize(2.0,2.0)</code> resizes to 200% of the original.<br>
     * <code>resize({@link Double#MAX_VALUE},{@link Double#MAX_VALUE})</code> resizes to the dimension of the embedded image.
     * </p>
     */
    private Double scaleX;
    private Double scaleY;

    public ImageCommand() {
    }

    /**
     * Creates the command from an image in the context
     * @param image name of the context attribute with the image bytes
     * @param imageType type of the image
     */
    public ImageCommand(String image, ImageType imageType) {
        this.src = image;
        this.imageType = imageType;
    }

    /**
     * Creates the command from the image bytes
     * @param imageBytes the image byte array
     * @param imageType the type of the image to render (e.g. PNG, JPEG etc)
     */
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

    public ImageType getImageType() {
        return imageType;
    }

    public void setImageType(ImageType imageType) {
        this.imageType = imageType;
    }

    /**
     * @param strType "PNG", "JPEG" (not "JPG"), ...
     */
    public void setImageType(String strType) {
        imageType = ImageType.valueOf(strType);
    }

    public Double getScaleX() {
        return scaleX;
    }

    public void setScaleX(String scaleX) {
        this.scaleX = Double.valueOf(scaleX);
    }

    public Double getScaleY() {
        return scaleY;
    }

    public void setScaleY(String scaleY) {
        this.scaleY = Double.valueOf(scaleY);
    }

    private boolean needResizePicture() {
        return this.scaleX != null && this.scaleY != null;
    }

    @Override
    public Boolean getLockRange() {
        return needResizePicture() ? Boolean.FALSE : super.getLockRange();
    }

    @Override
    public Command addArea(Area area) {
        if (areaList.size() >= 1) {
            throw new IllegalArgumentException("You can only add 1 area to 'image' command!");
        }
        this.area = area;
        return super.addArea(area);
    }

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        if (area == null) {
            throw new IllegalArgumentException("No area is defined for image command");
        }
        Size imageAnchorAreaSize = new Size(area.getSize().getWidth() + 1, area.getSize().getHeight() + 1);
        AreaRef imageAnchorArea = new AreaRef(cellRef, imageAnchorAreaSize);
        byte[] imgBytes = imageBytes;
        if (src != null) {
            Object imgObj = getTransformationConfig().getExpressionEvaluator().evaluate(src, context.toMap());
            if (imgObj == null) {
                return area.getSize();
            }
            if (!(imgObj instanceof byte[])) {
                throw new IllegalArgumentException("src value must contain image bytes (byte[])");
            }
            imgBytes = (byte[]) imgObj;
        }
        addImage(((PoiTransformer) getTransformer()).getWorkbook(), imageAnchorArea, imgBytes, imageType, scaleX, scaleY);
        return area.getSize();
    }
    
    private void addImage(Workbook workbook, AreaRef areaRef, byte[] imageBytes, ImageType imageType, Double scaleX, Double scaleY) {
        int poiPictureType = findPoiPictureTypeByImageType(imageType);
        int pictureIdx = workbook.addPicture(imageBytes, poiPictureType);
        addImage(workbook, areaRef, pictureIdx, scaleX, scaleY);
    }

    private int findPoiPictureTypeByImageType(ImageType imageType) {
        if (imageType == null) {
            throw new IllegalArgumentException("imageType must not be null");
        }
		switch (imageType) {
		case PNG:
			return Workbook.PICTURE_TYPE_PNG;
		case JPEG:
			return Workbook.PICTURE_TYPE_JPEG;
		case EMF:
			return Workbook.PICTURE_TYPE_EMF;
		case WMF:
			return Workbook.PICTURE_TYPE_WMF;
		case DIB:
			return Workbook.PICTURE_TYPE_DIB;
		case PICT:
			return Workbook.PICTURE_TYPE_PICT;
		default:
			return -1;
		}
    }

    private void addImage(Workbook workbook, AreaRef areaRef, int imageIdx, Double scaleX, Double scaleY) {
        boolean resize = scaleX != null && scaleY != null;
        CreationHelper helper = workbook.getCreationHelper();
        Sheet sheet = workbook.getSheet(areaRef.getSheetName());
        if (sheet == null) {
            sheet = workbook.createSheet(areaRef.getSheetName());
        }
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(areaRef.getFirstCellRef().getCol());
        anchor.setRow1(areaRef.getFirstCellRef().getRow());
        if (resize) {
            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
            anchor.setCol2(-1);
            anchor.setRow2(-1);
        } else {
            anchor.setCol2(areaRef.getLastCellRef().getCol());
            anchor.setRow2(areaRef.getLastCellRef().getRow());
        }
        Picture picture = drawing.createPicture(anchor, imageIdx);
        if (resize) {
            picture.resize(scaleX.doubleValue(), scaleY.doubleValue());
        }
    }
    
	/**
     * Reads all the data from the input stream and returns the bytes read.
     * 
     * @param inputStream -
     * @return byte array
     * @throws IOException -
     */
    public static byte[] toByteArray(InputStream inputStream) throws IOException { // used by templates and SimpleExporter
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[4096];
        int count;
        while ((count = inputStream.read(buffer)) != -1) {
            baos.write(buffer, 0, count);
        }
        return baos.toByteArray();
    }
}
