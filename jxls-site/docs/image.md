# Image

Use this command for adding an image to a sheet.

```
jx:image(src="image" lastCell="A2")
```

`src`: expression that returns a byte[] array that contains the image data

`imageType`: can have these values: PNG (default), JPEG (not JPG), EMF, WMF, PICT, DIB.

`scaleX` and `scaleY`: optional Double values for scaling

`lastCell`: area end

Here is example code for adding an image to the data map:

```
InputStream imageInputStream = ImageDemo.class.getResourceAsStream("business.png");
byte[] imageBytes = ImageCommand.toByteArray(imageInputStream);
data.put("image", imageBytes);
```

jx:image is part of jxls-poi and is only available if you use JxlsPoiTemplateFillerBuilder or add the ImageCommand using withCommand().

This command is a community contribution.
