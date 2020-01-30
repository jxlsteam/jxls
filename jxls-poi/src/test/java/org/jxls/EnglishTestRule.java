package org.jxls;

import java.util.Locale;

import org.junit.rules.TestRule;
import org.junit.runner.Description;
import org.junit.runners.model.Statement;

/**
 * Sets locale for testcase to ENGLISH.
 */
public class EnglishTestRule implements TestRule {

    @Override
    public Statement apply(final Statement base, Description description) {
        return new Statement() {
            @Override
            public void evaluate() throws Throwable {
                Locale before = Locale.getDefault();
                Locale.setDefault(Locale.ENGLISH);
                try {
                    base.evaluate();
                } finally {
                    Locale.setDefault(before); // restore
                }
            }
        };
    }
}
