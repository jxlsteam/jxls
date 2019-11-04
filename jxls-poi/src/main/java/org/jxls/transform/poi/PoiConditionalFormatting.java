package org.jxls.transform.poi;

import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class PoiConditionalFormatting {
    private List<ConditionalFormattingRule> rules = new ArrayList<>();
    private List<CellRangeAddress> ranges;

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

