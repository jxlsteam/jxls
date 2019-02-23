package org.jxls.builder.xml;

import ch.qos.logback.core.joran.GenericConfigurator;
import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.action.ImplicitAction;
import ch.qos.logback.core.joran.spi.ElementSelector;
import ch.qos.logback.core.joran.spi.Interpreter;
import ch.qos.logback.core.joran.spi.RuleStore;

import java.util.List;
import java.util.Map;

/**
 * Configurator used by XmlAreaBuilder to configure building rules
 * @author Leonid Vysochyn
 *         Date: 2/14/12
 */
class SimpleConfigurator extends GenericConfigurator {

    private final Map<ElementSelector, Action> ruleMap;
    private final List<ImplicitAction> iaList;

    SimpleConfigurator(Map<ElementSelector, Action> ruleMap) {
        this(ruleMap, null);
    }

    private SimpleConfigurator(Map<ElementSelector, Action> ruleMap, List<ImplicitAction> iaList) {
        this.ruleMap = ruleMap;
        this.iaList = iaList;
    }

    @Override
    protected void addInstanceRules(RuleStore rs) {
        for (ElementSelector pattern : ruleMap.keySet()) {
            Action action = ruleMap.get(pattern);
            rs.addRule(pattern, action);
        }
    }

    @Override
    protected void addImplicitRules(Interpreter interpreter) {
        if(iaList == null) {
            return;
        }
        for (ImplicitAction ia : iaList) {
            interpreter.addImplicitAction(ia);
        }
    }

}
