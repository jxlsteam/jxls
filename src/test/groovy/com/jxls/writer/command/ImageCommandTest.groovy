package com.jxls.writer.command

import spock.lang.Specification
import com.jxls.writer.area.Area
import com.jxls.writer.common.Context
import com.jxls.writer.common.CellRef
import com.jxls.writer.common.ImageType
import com.jxls.writer.transform.Transformer
import com.jxls.writer.common.Size
import com.jxls.writer.common.AreaRef

/**
 * @author Leonid Vysochyn
 * Date: 6/15/12 2:54 PM
 */
class ImageCommandTest extends Specification {
    def "test applyAt"(){
        given:
            def area = Mock(Area)
            def transformer = Mock(Transformer)
            def imgBytes = new byte[10]
            def imageCommand = new ImageCommand("image", ImageType.PNG)
            imageCommand.addArea(area)
            def context = new Context()
        when:
            context.putVar("image", imgBytes)
            imageCommand.applyAt(new CellRef(5, 5), context)
        then:
            area.getSize() >> new Size(3,4)
            1 * area.getTransformer() >> transformer
            1 * transformer.addImage(new AreaRef(new CellRef(5,5), new Size(3,4)), imgBytes, ImageType.PNG)
            0 * _._
    }
}
