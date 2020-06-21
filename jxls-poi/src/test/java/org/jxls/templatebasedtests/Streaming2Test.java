package org.jxls.templatebasedtests;

import org.jxls.common.Context;

/**
 * Sheet 1 with Streaming, sheet 2 without streaming and formulas
 */
public class Streaming2Test extends StreamingTest {

    @Override
    protected Context createTestData() {
        Context ctx = super.createTestData();
        ctx.putVar("discount", Double.valueOf(5));
        return ctx;
    }
}
