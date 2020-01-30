package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xml.XmlAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 */
public class ObjectCollectionXMLBuilderDemo {
    private static final Logger logger = LoggerFactory.getLogger(ObjectCollectionXMLBuilderDemo.class);

    @Test
    public void test() throws ParseException, IOException {
        logger.info("Running Object Collection XML Builder demo");
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try (InputStream is = ObjectCollectionXMLBuilderDemo.class.getResourceAsStream("object_collection_xmlbuilder_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/object_collection_xmlbuilder_output.xls")) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                try (InputStream configInputStream = ObjectCollectionXMLBuilderDemo.class.getResourceAsStream("object_collection_xmlbuilder.xml")) {
                    AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
                    List<Area> xlsAreaList = areaBuilder.build();
                    Area xlsArea = xlsAreaList.get(0);
                    Context context = new Context();
                    context.putVar("employees", employees);
                    xlsArea.applyAt(new CellRef("Result!A1"), context);
                    transformer.write();
                }
            }
        }
    }
}
