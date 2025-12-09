# Tinylog <!-- ** -->

## Version 1

```
implementation 'org.tinylog:tinylog:1.3.6'
```

Add this class to your project and also add the above dependency to it. Then you can configure [Tinylog](https://tinylog.org/v1/)
and use it with Jxls.

```
import org.jxls.common.PoiExceptionLogger;
import org.pmw.tinylog.Logger;

public class TinylogJxlsLogger extends PoiExceptionLogger {

    @Override
    public void debug(String msg) {
        Logger.debug(msg);
    }

    @Override
    public void info(String msg) {
        Logger.info(msg);
    }
    
    @Override
    public void warn(String msg) {
        Logger.warn(msg);
    }
    
    @Override
    public void warn(Throwable e, String msg) {
        Logger.warn(e, msg);
    }
    
    @Override
    public void error(String msg) {
        Logger.error(msg);
    }
    
    @Override
    public void error(Throwable e, String msg) {
        Logger.error(e, msg);
    }
}
```

Use that class with Jxls:

```
// configure Logger ...
JxlsPoiTemplateFillerBuilder.newInstance().withLogger(new TinylogJxlsLogger())
```
