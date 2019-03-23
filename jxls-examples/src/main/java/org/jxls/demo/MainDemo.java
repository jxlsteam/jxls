package org.jxls.demo;

import org.jxls.demo.guide.GroupingDemo;
import org.jxls.demo.guide.HighlightDemo;
import org.jxls.demo.guide.IfCommandDemo;
import org.jxls.demo.guide.MultiSheetDemo;
import org.jxls.demo.guide.MultiSheetPropertyExpressionDemo;
import org.jxls.demo.guide.NestedCommandDemo;
import org.jxls.demo.guide.NestedCommandJavaAPIDemo;
import org.jxls.demo.guide.ObjectCollectionByColumnsDemo;
import org.jxls.demo.guide.ObjectCollectionDemo;
import org.jxls.demo.guide.ObjectCollectionFormulasDemo;
import org.jxls.demo.guide.ObjectCollectionJavaAPIDemo;
import org.jxls.demo.guide.ObjectCollectionXMLBuilderDemo;
import org.jxls.demo.guide.ParameterizedFormulasDemo;
import org.jxls.demo.reader.XlsReaderDemo;
import org.jxls.util.TransformerFactory;

import java.io.File;

/**
 * @author Leonid Vysochyn
 *         Date: 2/9/12 5:05 PM
 */
public class MainDemo {
    public static void main(String[] args) throws Exception {
        File targetDir = new File("target");
        if( !targetDir.exists() ){
            targetDir.mkdir();
        }
        ObjectCollectionDemo.main(args);
        ObjectCollectionByColumnsDemo.main(args);
        ObjectCollectionJavaAPIDemo.main(args);
        ObjectCollectionFormulasDemo.main(args);
        ParameterizedFormulasDemo.main(args);
        ObjectCollectionXMLBuilderDemo.main(args);
        NestedCommandJavaAPIDemo.main(args);
        NestedCommandDemo.main(args);
        HighlightDemo.main(args);
        IfCommandDemo.main(args);
        GroupingDemo.main(args);
        MultiSheetDemo.main(args);
        MultiSheetPropertyExpressionDemo.main(args);

        EachIfCommandDemo.main(args);
        EachIfXmlBuilderDemo.main(args);
        FormulaExportDemo.main(args);

        MultipleSheetDemo.main(args);
        XlsCommentBuilderDemo.main(args);
        ImageDemo.main(args);
        DynamicColumnsDemo.main(args);
        GridCommandDemo.main(args);

        SimpleExporterDemo.main(args);

        FormulaCopyDemo.main(args);

        TwoInnerLoopsDemo.main(args);

        XlsReaderDemo.main(args);

        MultiSheetMarkupDemo.main(args);
        MultiSheetMarkupDemo2.main(args);
        JexlCustomFunctionDemo.main(args);
        CustomExpressionNotationDemo.main(args);
        ShiftStrategyDemo.main(args);

        ShiftWithEmptyList.main(args);
        SqlDemo.main(args);

        String transformerName = TransformerFactory.getTransformerName();

        if( TransformerFactory.POI_TRANSFORMER.equals( transformerName ) ){
            UserCommandExcelMarkupDemo.main(args);
            UserCommandDemo.main(args);
            AreaListenerDemo.main(args);
            StressXlsxDemo.executeStress1();
            StressXlsxDemo.executeStress2();
            SxssfDemo.simpleSxssf();
            SxssfDemo.executeStress1();
            SxssfDemo.executeStress2();
        }

        if( TransformerFactory.JEXCEL_TRANSFORMER.equals( transformerName)){
            JexcelAreaListenerDemo.main(args);
            JexcelUserCommandExcelMarkupDemo.main(args);
        }

//        StressDemo.executeStress1();
//        StressDemo.executeStress2();
    }
}
