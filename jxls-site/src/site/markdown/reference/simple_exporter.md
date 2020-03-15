SimpleExporter
=================

Introduction
------------
You can export a list of objects into Excel with a single line of code by using *SimpleExporter* class.

This is achieved by using a special [built-in template](../xls/builtin_template.xls) containing [GridCommand](grid_command.html).

How to use
----------
Just create a *SimpleExporter* instance and run its `gridExport` method

    new SimpleExporter().gridExport(headers, dataObjects, propertyNames, outputStream);

Where 

* `headers` - a collection of headers
* `dataObjects` - a collection of data objects
* `propertyNames` - a comma-separated list of object properties
* `outputStream` - an output stream to write the final Excel      

See [SimpleExporter example](../samples/simple_exporter.html) to see it in action.

Custom Template
-----------------
You may register your own template to be used by *SimpleExporter* with its *registerGridTemplate* method.

    public void registerGridTemplate(InputStream templateInputStream)

The template must have a [GridCommand](grid_command.html) defined in it.

See [SimpleExporter example](../samples/simple_exporter.html) for an example on how to do it.

Transformer Support Note
-------------------------
Since [GridCommand](grid_command.html is currently supported only in POI transformer you have to use POI when working with *SimpleExporter*.