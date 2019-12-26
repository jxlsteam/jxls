package org.jxls.examples;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.JxlsTester.TransformerChecker;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.transform.Transformer;

public class CustomExpressionNotationDemo {

    @Test
    public void test() {
        Context context = new Context();
        context.putVar("employees", Employee.generateSampleEmployeeData());

        TransformerChecker useOtherExpressionNotations = new TransformerChecker() {
            @Override
            public Transformer checkTransformer(Transformer transformer) {
                transformer.getTransformationConfig().buildExpressionNotation("[[", "]]");
                return transformer;
            }
        };                
        
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.createTransformerAndProcessTemplate(context, useOtherExpressionNotations);
    }
}
