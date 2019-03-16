import static ch.qos.logback.classic.Level.WARN
import static ch.qos.logback.classic.Level.DEBUG
import ch.qos.logback.core.ConsoleAppender
import ch.qos.logback.classic.encoder.PatternLayoutEncoder

appender("CONSOLE", ConsoleAppender) {
    encoder(PatternLayoutEncoder) {
        pattern = "%level %logger - %msg%n"
    }
}

root(WARN, ['CONSOLE'])

logger("org.jxls.writer", DEBUG, ["CONSOLE"])