package org.jxls.template

import org.jxls.util.TransformerFactory
import spock.lang.Specification

/**
 * Created by Leonid Vysochyn on 19-Jul-15.
 */
class SimpleExporterTest extends Specification{
    // TODO
    def "test gridExport"(){
        when:
        def simpleExport = new SimpleExporter()
        def os = Mock(OutputStream)
        def transformerFactory = Mock(TransformerFactory) // mocking of java static methods probably is not supported

        def obj1 = ["name": "Alisa", "age": 30, "amount": 125.50]
        def obj2 = ["name": "Sergey", "age": 32, "amount": 150.40]
        def obj3 = ["name": "Olga", "age": 30, "amount": 200.15]
        def people = [obj1, obj2, obj3]
//        simpleExport.gridExport(["Name, Age, Amount"], people, "name, age, amount", os)
        then:
            1==1
    }
}
