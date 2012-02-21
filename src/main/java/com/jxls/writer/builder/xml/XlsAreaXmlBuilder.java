package com.jxls.writer.builder.xml;

import ch.qos.logback.core.Context;
import ch.qos.logback.core.ContextBase;
import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.action.NewRuleAction;
import ch.qos.logback.core.joran.spi.JoranException;
import ch.qos.logback.core.joran.spi.Pattern;
import ch.qos.logback.core.util.StatusPrinter;
import com.jxls.writer.area.Area;
import com.jxls.writer.builder.AreaBuilder;
import com.jxls.writer.transform.Transformer;

import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 *         Date: 2/14/12 11:50 AM
 */
public class XlsAreaXmlBuilder implements AreaBuilder{
    Transformer transformer;

    public XlsAreaXmlBuilder(Transformer transformer) {
        this.transformer = transformer;
    }

    public List<Area> build(InputStream is) {
        Map<Pattern, Action> ruleMap = new HashMap<Pattern, Action>();

        AreaAction areaAction = new AreaAction(transformer);
        ruleMap.put(new Pattern("*/area"), areaAction);
        ruleMap.put(new Pattern("*/each"), new EachAction());
        ruleMap.put(new Pattern("*/if"), new IfAction());
        ruleMap.put(new Pattern("*/user-command"), new UserCommandAction());

        ruleMap.put(new Pattern("*/user-action"), new NewRuleAction());

        Context context = new ContextBase();
        SimpleConfigurator simpleConfigurator = new SimpleConfigurator(ruleMap);
        simpleConfigurator.setContext(context);

        try {
            simpleConfigurator.doConfigure(is);
        } catch (JoranException e) {
            printOccurredErrors(context);
        }
        return areaAction.getAreaList();
    }

    private void printOccurredErrors(Context context) {
        StatusPrinter.print(context);
    }
}
