package org.jxls.common

import org.jxls.command.AbstractCommand
import org.jxls.expression.JexlExpressionEvaluator
import org.jxls.transform.AbstractTransformer
import org.jxls.transform.TransformationConfig
import org.jxls.transform.Transformer
import spock.lang.Specification

/**
 * @author Leonid Vysochyn
 * Date: 2/3/12 1:19 PM
 */
class CellDataTest extends Specification{
    def "test creation with (sheetName, col, row) params"(){
        when:
            CellData cellData = new CellData("sheet1", 2, 3)
        then:
            "sheet1" == cellData.getSheetName()
            2 == cellData.getRow()
            3 == cellData.getCol()
    }

    def "test equality"(){
        expect:
            new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc").equals( new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc"))
            !new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc1").equals( new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc"))
    }
    
    def "test creation with (sheetName, row, col, type, value)"(){
        when:
            CellData cellData = new CellData("Sheet1", 5, 10, CellData.CellType.NUMBER, 1.2)
        then:
            "Sheet1" == cellData.getSheetName()
            5 == cellData.getRow()
            10 == cellData.getCol()
            CellData.CellType.NUMBER == cellData.getCellType()
            1.2 == cellData.getCellValue()
            new CellRef("Sheet1",5,10) == cellData.getCellRef()
    }
    
    def "test creation with (pos, type, value)"(){
        when: 
            CellData cellData = new CellData(new CellRef("Sheet1",5,10), CellData.CellType.STRING, "Abc")
        then:
            new CellRef("Sheet1",5,10) == cellData.getCellRef()
            "Sheet1" == cellData.getSheetName()
            5 == cellData.getRow()
            10 == cellData.getCol()
    }


    def "test add target pos"(){
        when:
            CellData cellData = new CellData("sheet1", 2, 3)
            cellData.addTargetPos(new CellRef("sheet1",2,3))
            cellData.addTargetPos(new CellRef("sheet2",3,4))
            def targetPos = cellData.getTargetPos()
        then:
            targetPos.size() == 2
            targetPos.contains(new CellRef("sheet1", 2,3))
            targetPos.contains(new CellRef("sheet2",3,4))
    }    
    
    def "test get pos"(){
        when:
            CellData cellData = new CellData("sheet1",2,3)
        then:
            cellData.getCellRef() == new CellRef("sheet1",2,3)
    }
    
    def "test create with formula value"(){
        expect:
            new CellData("sheet1", 1, 2, cellType, formula).getFormula() == formulaValue
        where:
            formula             | formulaValue      | cellType
            "SUM(A1:A3)+B10"    | "SUM(A1:A3)+B10"  | CellData.CellType.FORMULA
            '$[SUM(A3:A7)+B12]' | "SUM(A3:A7)+B12"  | CellData.CellType.STRING
    }
    
    def "test reset target pos"(){
        when:
            def cellData = new CellData("sheet1",1,2, CellData.CellType.STRING, "Abc")
            cellData.addTargetPos(new CellRef("sheet2",0,0))
            cellData.addTargetPos(new CellRef("sheet1",1,1))
            cellData.resetTargetPos()
        then:
            cellData.getTargetPos().isEmpty()
    }

    def "test set/get comment"(){
        when:
            def cellData = new CellData("sheet1", 1,2, CellData.CellType.STRING, "Abc")
            cellData.setCellComment("Test comment")
        then:
            cellData.getCellComment() == "Test comment"
    }

    def "test evaluate number"(){
        setup:
            def cellData = new CellData("sheet1", 1, 2, CellData.CellType.STRING, '${num}')
            Integer num = new Integer(100)
            def context = new Context()
            context.putVar("num", num)
            def transformer = Mock(Transformer);
            def transformationConfig = new TransformationConfig()
            cellData.setTransformer(transformer)
        when:
            transformer.getTransformationConfig() >> transformationConfig
            def result = cellData.evaluate(context)
        then:
            cellData.targetCellType == CellData.CellType.NUMBER
            result == 100
    }

    def "test evaluate boolean"(){
        setup:
            def cellData = new CellData("sheet1", 1, 2, CellData.CellType.STRING, '${flag}')
            boolean flag = true
            def context = new Context()
            context.putVar("flag", flag)
            def transformer = Mock(Transformer);
            def transformationConfig = new TransformationConfig()
            cellData.setTransformer(transformer)
        when:
            transformer.getTransformationConfig() >> transformationConfig
            def result = cellData.evaluate(context)
        then:
            cellData.targetCellType == CellData.CellType.BOOLEAN
            result == true
    }

    def "test evaluate date"(){
        setup:
            def cellData = new CellData("sheet1", 1, 2, CellData.CellType.STRING, '${today}')
            Date today = new Date()
            def context = new Context()
            context.putVar("today", today)
            def transformer = Mock(Transformer);
            def transformationConfig = new TransformationConfig()
            cellData.setTransformer(transformer)
        when:
            transformer.getTransformationConfig() >> transformationConfig
            def result = cellData.evaluate(context)
        then:
            cellData.targetCellType == CellData.CellType.DATE
            result == today
    }

    def "test evaluate user formula"(){
        setup:
            def cellData = new CellData("sheet1", 1, 2, CellData.CellType.STRING, '$[SUM(B2:B4)* (1 + ${bonus})]')
            def context = new Context()
            context.putVar("bonus", 0.15)
            def transformer = Mock(Transformer);
            def transformationConfig = new TransformationConfig()
            cellData.setTransformer(transformer)
        when:
            transformer.getTransformationConfig() >> transformationConfig
            def result = cellData.evaluate(context)
        then:
            cellData.targetCellType == CellData.CellType.FORMULA
            result == "SUM(B2:B4)* (1 + 0.15)"
    }

    def "test evaluate combined expression"(){
        setup:
            def cellData = new CellData("sheet1", 1, 2, CellData.CellType.STRING, '${num} days')
            Integer num = new Integer(21)
            def context = new Context()
            context.putVar("num", num)
            def transformer = Mock(Transformer);
            def transformationConfig = new TransformationConfig()
        when:
            cellData.setTransformer(transformer)
            transformer.getTransformationConfig() >> transformationConfig
            def result = cellData.evaluate(context)
        then:
            cellData.targetCellType == CellData.CellType.STRING
            result == "21 days"
    }

    def "test evaluate another combined expression"(){
        setup:
            def cellData = new CellData("sheet1", 1, 2, CellData.CellType.STRING, 'Days: ${num}')
            Integer num = new Integer(21)
            def context = new Context()
            context.putVar("num", num)
            def transformer = Mock(Transformer);
            def transformationConfig = new TransformationConfig()
        when:
            cellData.setTransformer(transformer)
            transformer.getTransformationConfig() >> transformationConfig
            def result = cellData.evaluate(context)
        then:
            cellData.targetCellType == CellData.CellType.STRING
            result == "Days: 21"
    }

}
