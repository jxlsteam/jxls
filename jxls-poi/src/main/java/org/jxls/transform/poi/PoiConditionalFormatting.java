package org.jxls.transform.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

class PoiConditionalFormatting {
    List<ConditionalFormattingRule> rules = new ArrayList<>();
    List<CellRangeAddress> ranges;

    PoiConditionalFormatting(ConditionalFormatting conditionalFormatting) {
        for(int i = 0; i < conditionalFormatting.getNumberOfRules(); i++){
            rules.add(conditionalFormatting.getRule(i));
        }
        ranges = Arrays.asList(conditionalFormatting.getFormattingRanges());
    }
}

