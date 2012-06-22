package com.jxls.plus.command

import spock.lang.Specification
import com.jxls.plus.area.Area
import com.jxls.plus.common.Context
import com.jxls.plus.common.CellRef
import com.jxls.plus.common.ImageType
import com.jxls.plus.transform.Transformer
import com.jxls.plus.common.Size
import com.jxls.plus.common.AreaRef
import com.jxls.plus.command.ImageCommand

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
