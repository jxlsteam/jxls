package org.jxls.command

import org.jxls.transform.TransformationConfig
import org.jxls.transform.Transformer
import spock.lang.Specification
import org.jxls.common.CellRef
import org.jxls.area.Area
import org.jxls.common.Context

/**
 * @author Leonid Vysochyn
 * Date: 1/5/12
 */
class IfCommandTest extends Specification{
    def "test init" (){
        when:
            def ifArea = Mock(Area)
            def elseArea = Mock(Area)
            def ifCommand = new IfCommand("2*x + 5 > 10",
                    ifArea, elseArea );
        then:
             ifCommand.condition == "2*x + 5 > 10"
             ifCommand.ifArea == ifArea
             ifCommand.elseArea == elseArea
    }
    
    def "test add area"(){
        def ifArea = Mock(Area)
        def elseArea = Mock(Area)
        when:
            def ifCommand = new IfCommand("a > b")
            ifCommand.addArea(ifArea)
            ifCommand.addArea(elseArea)
        then:
            ifCommand.condition == "a > b"
            ifCommand.ifArea == ifArea
            ifCommand.elseArea == elseArea
    }

    def "test add excessive number of areas"(){
        def ifCommand = new IfCommand("a > b")
        ifCommand.addArea(Mock(Area))
        ifCommand.addArea(Mock(Area))
        when:
            ifCommand.addArea(Mock(Area))
        then:
            thrown(IllegalArgumentException)
    }

    def "test applyAt when condition is false"(){
        given:
            def ifArea = Mock(Area)
            def elseArea = Mock(Area)
            def transformer = Mock(Transformer)
            def transformationConfig = new TransformationConfig()
            def ifCommand = new IfCommand("2*x + 5 > 10", ifArea, elseArea)
            def context = new Context()
        when:
            context.putVar("x", 2)
            ifCommand.applyAt(new CellRef(1, 1), context)
        then:
            1 * elseArea.applyAt(new CellRef(1, 1), context)
            1 * ifArea.getTransformer() >> transformer
            1 * transformer.getTransformationConfig() >> transformationConfig
            0 * _._
    }

    def "test applyAt when condition is true"(){
        given:
            def ifArea = Mock(Area)
            def elseArea = Mock(Area)
            def transformer = Mock(Transformer)
            def transformationConfig = new TransformationConfig()
            def ifCommand = new IfCommand("2*x + 5 > 10", ifArea, elseArea)
            def context = new Context()
        when:
            context.putVar("x", 3)
            ifCommand.applyAt(new CellRef(1, 1), context)
        then:
            1 * ifArea.applyAt(new CellRef(1, 1), context)
            1 * ifArea.getTransformer() >> transformer
            1 * transformer.getTransformationConfig() >> transformationConfig
            0 * _._
    }

}
