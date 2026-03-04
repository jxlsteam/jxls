# SLF4J <!-- ** -->

**Simple Logging Facade for Java**

```
implementation group: 'org.slf4j', name: 'slf4j-api', version: '2.0.6'
```

Add this class to your project and also add the above dependency to it. Then you can configure [SLF4J](https://www.slf4j.org)
and use it with Jxls. Because slf4j-api is just the facade, another dependency needs to be added with the implementation.

```
import org.jxls.common.PoiExceptionLogger;
import org.slf4j.Logger;

public class Slf4jJxlsLogger extends PoiExceptionLogger {
    private final Logger logger;
    
    public Slf4jJxlsLogger(Logger logger) {
        this.logger = logger;
    }

    @Override
    public void debug(String msg) {
        logger.debug(msg);
    }

    @Override
    public void info(String msg) {
        logger.info(msg);
    }

    @Override
    public void warn(String msg) {
        logger.warn(msg);
    }

    @Override
    public void warn(Throwable e, String msg) {
        logger.warn(msg, e);
    }

    @Override
    public void error(String msg) {
        logger.error(msg);
    }

    @Override
    public void error(Throwable e, String msg) {
        logger.error(msg, e);
    }
}
```

Use that class with Jxls:

```
Logger logger = ...
JxlsPoiTemplateFillerBuilder.newInstance().withLogger(new Slf4jJxlsLogger(logger))
```
