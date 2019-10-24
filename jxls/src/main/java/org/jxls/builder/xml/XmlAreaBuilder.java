package org.jxls.builder.xml;

import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.builder.AreaBuilder;
import org.jxls.transform.Transformer;

import ch.qos.logback.core.Context;
import ch.qos.logback.core.ContextBase;
import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.action.NewRuleAction;
import ch.qos.logback.core.joran.spi.ElementSelector;
import ch.qos.logback.core.joran.spi.JoranException;
import ch.qos.logback.core.util.StatusPrinter;

/**
 * Creates an area based on the XML definition
 * 
 * @author Leonid Vysochyn
 */
public class XmlAreaBuilder implements AreaBuilder {
    private final InputStream xmlInputStream;
    private Transformer transformer;
    private final boolean clearTemplateCells;

    public XmlAreaBuilder(Transformer transformer) {
        this(null, transformer, true);
    }

    public XmlAreaBuilder(InputStream xmlInputStream, Transformer transformer) {
        this(xmlInputStream, transformer, true);
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
        // TODO ImageAction
        // TODO UpdateCellAction
        // TODO MergeCellsAction (plus documentation)

        ruleMap.put(new ElementSelector("*/user-action"), new NewRuleAction());

        Context context = new ContextBase();
        SimpleConfigurator simpleConfigurator = new SimpleConfigurator(ruleMap);
        simpleConfigurator.setContext(context);

        try {
            simpleConfigurator.doConfigure(is);
        } catch (JoranException e) {
            printOccurredErrors(context);
        }
        if (clearTemplateCells) {
            for (Area area : areaAction.getAreaList()) {
                ((XlsArea) area).clearCells();
            }
        }
        return areaAction.getAreaList();
    }

    @Override
    public List<Area> build() {
        return build(xmlInputStream);
    }

    private void printOccurredErrors(Context context) {
        StatusPrinter.print(context);
    }
}
