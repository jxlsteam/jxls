package org.jxls;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;

/**
 * Template based test classes use this class for processing the template.
 */
public class JxlsTester implements AutoCloseable {
    private final Class<?> testclass;
    private final String excelTemplateFilename;
    private final File out;
    private boolean useFastFormulaProcessor = false;
    /** evaluating formulas turned on by default because in testcases we want to verify the output files */
    private boolean evaluateFormulas = true;

    /**
     * Use this constructor if you really need to change the Excel template filename (reasons can be: different templates in one testclass;
     * XLS format) or you need to extend this class.
     * @param testclass
     * @param excelTemplateFilename
     */
    private JxlsTester(Class<?> testclass, String excelTemplateFilename) {
        this.testclass = testclass;
        this.excelTemplateFilename = excelTemplateFilename;
        int o = excelTemplateFilename.lastIndexOf(".");
        File folder = new File("target");
        folder.mkdir();
        out = new File(folder, excelTemplateFilename.substring(0, o) + "_output" + excelTemplateFilename.substring(o));
    }

    /**
     * Most common way to create instance. XLSX format.
     * @param testclass test class name results in Excel template filename + ".xlsx"
     * @return new JxlsTester instance
     */
    public static JxlsTester xlsx(Class<?> testclass) {
        return new JxlsTester(testclass, testclass.getSimpleName() + ".xlsx");
    }

    /**
     * XLS format
     * @param testclass test class name results in Excel template filename + ".xls"
     * @return new JxlsTester instance
     */
    public static JxlsTester xls(Class<?> testclass) {
        return new JxlsTester(testclass, testclass.getSimpleName() + ".xls");
    }

    /**
     * Excel template filename will be build from testclass name + "_" + method + ".xlsx".
     * @param testclass Java test class name
     * @param method    Java test method name
     * @return new JxlsTester instance
     */
    public static JxlsTester xlsx(Class<?> testclass, String method) {
        return new JxlsTester(testclass, testclass.getSimpleName() + "_" + method + ".xlsx");
    }
    
    /**
     * Most common method for processing the template.
     * @param context context for processing template
     */
    public void processTemplate(Context context) {
        try (InputStream is = testclass.getResourceAsStream(excelTemplateFilename)) {
            try (OutputStream os = new FileOutputStream(out)) {
                JxlsHelper jxls = JxlsHelper.getInstance();
                if (useFastFormulaProcessor) {
                    jxls.setUseFastFormulaProcessor(useFastFormulaProcessor);
                }
                jxls.setEvaluateFormulas(evaluateFormulas);
                jxls.processTemplate(is, os, context);
            }
        } catch (IOException e) { // Testcase does not need not catch IOException.
            throw new RuntimeException(e);
        }
    }

    /**
     * Call this method if you need to access the transformer instance.
     * @param context context for processing template
     * @param transformerChecker object for checking the transformer
     */
    public void createTransformerAndProcessTemplate(Context context, TransformerChecker transformerChecker) {
        try (InputStream is = testclass.getResourceAsStream(excelTemplateFilename)) {
            try (OutputStream os = new FileOutputStream(out)) {
                Transformer transformer = PoiTransformer.createTransformer(is, os);
                transformer = transformerChecker.checkTransformer(transformer);
                JxlsHelper jxls = JxlsHelper.getInstance();
                if (useFastFormulaProcessor) {
                    jxls.setUseFastFormulaProcessor(useFastFormulaProcessor);
                }
                jxls.setEvaluateFormulas(evaluateFormulas);
                jxls.processTemplate(context, transformer);
            }
        } catch (IOException e) { // Testcase does not need not catch IOException.
            throw new RuntimeException(e);
        }
    }
    
    // [Java 8] Uses can be changed to use of lambda
    /**
     * Check, edit or change transformer instance
     */
    public interface TransformerChecker {
        
        /**
         * @param transformer
         * @return usually transformer
         */
        Transformer checkTransformer(Transformer transformer);
    }

    /**
     * Use with try notation!
     * @return TestWorkbook for verifying the created Excel file
     */
    public TestWorkbook getWorkbook() {
        return new TestWorkbook(out);
    }

    /** We don't care if the output file won't be deleted. The use of the try notation is not mandatory. */
    @Override
    public void close() throws Exception {
        out.delete();
    }
    
    /**
     * Special method just for getting the input stream
     * @param testclass
     * @param xlsx true: XLSX format, false: XLS format
     * @return new opened InputStream
     * @throws IOException if file not found
     */
    public static InputStream openInputStream(Class<?> testclass, boolean xlsx) throws IOException {
        return testclass.getResourceAsStream(testclass.getSimpleName() + (xlsx ? ".xlsx" : ".xls"));
    }

    public JxlsTester dontEvaluateFormulas() {
        evaluateFormulas = false;
        return this;
    }

    public boolean isUseFastFormulaProcessor() {
        return useFastFormulaProcessor;
    }

    public void setUseFastFormulaProcessor(boolean useFastFormulaProcessor) {
        this.useFastFormulaProcessor = useFastFormulaProcessor;
    }

    public String getTemplateFilename() {
        return excelTemplateFilename;
    }
    
    public static void quickProcessXlsTemplate(Class<?> testcase) {
        Context context = new Context();
        context.putVar("employees", Employee.generateSampleEmployeeData());

        JxlsTester tester = JxlsTester.xls(testcase);
        tester.processTemplate(context);
    }
}
