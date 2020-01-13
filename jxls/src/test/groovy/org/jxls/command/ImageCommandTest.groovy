package org.jxls.command

import org.jxls.transform.TransformationConfig
import spock.lang.Specification
import org.jxls.area.Area
import org.jxls.common.Context
import org.jxls.common.CellRef
import org.jxls.common.ImageType
import org.jxls.transform.Transformer
import org.jxls.common.Size
import org.jxls.common.AreaRef

/**
 * @author Leonid Vysochyn
 * Date: 6/15/12 2:54 PM
 */
class ImageCommandTest extends Specification {
    def "test applyAt"(){
        given:
            def area = Mock(Area)
            def transformer = Mock(Transformer)
            def transformationConfig = new TransformationConfig()
            def imgBytes = new byte[10]
            def imageCommand = new ImageCommand("image", ImageType.PNG)
            imageCommand.addArea(area)
            def context = new Context()
        when:
            context.putVar("image", imgBytes)
            imageCommand.applyAt(new CellRef(5, 5), context)
        then:
            area.getSize() >> new Size(3,4)
            2 * area.getTransformer() >> transformer
            1 * transformer.addImage(new AreaRef(new CellRef(5,5), new Size(4,5)), imgBytes, ImageType.PNG, _, _)
            1 * transformer.getTransformationConfig() >> transformationConfig
            0 * _._
    }
}
