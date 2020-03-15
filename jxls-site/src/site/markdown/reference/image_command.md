Image-Command
==============

Introduction
------------

*Image-command* is used to output an image in Excel report.

Usage
-----
Have a look into ImageDemo.java to see *Image-command* in action.
The example demonstrates how to define the *Image-command* with Java API.
 
If you wish to use Excel mark-up to define the command it will look like this

    jx:image(lastCell="D10" src="image" imageType="PNG")
    
Here *lastCell* defines the right bottom cell of the image containing area. 
If the comment is placed for example in cell `A1` then the image will be placed at `A1:D10` area in the output Excel.
*src* defines the bean name in the jxls Context containing the image bytes.
*imageType* defines the image type and can be one of the following - *PNG, JPEG, EMF, WMF, PICT, DIB*.
The default value for *imageType* attribute is *PNG*. So in the above example we could skip it.

The Java code should contain something like this to put the required image into the Context under image attribute

    InputStream imageInputStream = ImageDemo.class.getResourceAsStream("business.png");
    byte[] imageBytes = Util.toByteArray(imageInputStream);
    context.putVar("image", imageBytes);
    