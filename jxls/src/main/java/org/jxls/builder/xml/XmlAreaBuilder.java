package org.jxls.builder.xml;

import ch.qos.logback.core.Context;
import ch.qos.logback.core.ContextBase;
import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.action.NewRuleAction;
import ch.qos.logback.core.joran.spi.ElementSelector;
import ch.qos.logback.core.joran.spi.JoranException;
import ch.qos.logback.core.util.StatusPrinter;
import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.builder.AreaBuilder;
import org.jxls.transform.Transformer;

import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Creates an area through  XML definition
 * @author Leonid Vysochyn
 *         Date: 2/14/12 11:50 AM
 */
public class XmlAreaBuilder implements AreaBuilder {
    private Transformer transformer;
    private InputStream xmlInputStream;
    private boolean clearTemplateCells = true;

    public XmlAreaBuilder(Transformer transformer) {
        this.transformer = transformer;
    }

    public XmlAreaBuilder(InputStream xmlInputStream, Transformer transformer) {
        this.xmlInputStream = xmlInputStream;
        this.transformer = transformer;
    }

    public XmlAreaBuilder(InputStream xmlInputStream, Transformer transformer, boolean clearTemplateCells) {
        this.xmlInputStream = xmlInputStream;
        this.transformer = transformer;
        this.clearTemplateCells = clearTemplateCells;
    }

    @Override
    public Transformer getTransformer() {
        return transformer;
    }

    @Override
    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }

    public List<Area> build(InputStream is) {
        Map<ElementSelector, Action> ruleMap = new HashMap<>();

        AreaAction areaAction = new AreaAction(transformer);
        ruleMap.put(new ElementSelector("*/area"), areaAction);
        ruleMap.put(new ElementSelector("*/each"), new EachAction());
        ruleMap.put(new ElementSelector("*/if"), new IfAction());
        ruleMap.put(new ElementSelector("*/user-command"), new UserCommandAction());
        ruleMap.put(new ElementSelector("*/grid"), new GridAction());

        ruleMap.put(new ElementSelector("*/user-action"), new NewRuleAction());

        Context context = new ContextBase();
        SimpleConfigurator simpleConfigurator = new SimpleConfigurator(ruleMap);
        simpleConfigurator.setContext(context);

        try {
            simpleConfigurator.doConfigure(is);
        } catch (JoranException e) {
            printOccurredErrors(context);
        }
        if( clearTemplateCells ){
            for(Area area: areaAction.getAreaList()){
                ((XlsArea)area).clearCells();
            }
        }
        return areaAction.getAreaList();
    }

    public List<Area> build() {
        return build(xmlInputStream);
    }

    private void printOccurredErrors(Context context) {
        StatusPrinter.print(context);
    }
}
