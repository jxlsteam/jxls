package org.jxls.transform.poi;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.util.CellRangeAddress;

public class PoiConditionalFormatting {
    private final List<ConditionalFormattingRule> rules = new ArrayList<>();
    private final List<CellRangeAddress> ranges;

    PoiConditionalFormatting(ConditionalFormatting conditionalFormatting) {
        for(int i = 0; i < conditionalFormatting.getNumberOfRules(); i++){
            rules.add(conditionalFormatting.getRule(i));
        }
        ranges = Arrays.asList(conditionalFormatting.getFormattingRanges());
    }

    public List<ConditionalFormattingRule> getRules() {
        return rules;
    }

    public List<CellRangeAddress> getRanges() {
        return ranges;
    }
}
