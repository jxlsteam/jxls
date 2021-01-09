package org.jxls.templatebasedtests;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.transform.poi.SelectSheetsForStreamingPoiTransformer;
import org.jxls.transform.poi.WritableCellValue;
import org.jxls.util.JxlsHelper;

public class Issue81Test {
    // TODO change class to standard testcase style and add verifications
    private Map<String, Object> model;
    
    @Before
    public void init() {
        model = new HashMap<>();
        model.put("list", createList());
    }
    
    @Test
    public void withStreaming() throws Exception {
        exportExcel(model, "/org/jxls/templatebasedtests/Issue81Test.xlsx", "target/Issue81Test_output.true.xlsx", true);
        // error: red cells are missing
        /* Liao Xue Wei: One is that the following columns of using direction="RIGHT" are covered; The other is, the
         * unit:mergeCell extension function will cause the contents of the previous column header to be cleared.
         * If there is no excel comment with unit:mergeCell, the previous column header will output normally.*/
    }
    
    @Test
    public void noStreaming() throws Exception {
        exportExcel(model, "/org/jxls/templatebasedtests/Issue81Test.xlsx", "target/Issue81Test_output.false.xlsx", false);
        // Result is ok
    }

    private void exportExcel(Map<String, Object> model, String inResourceFileName, String outFileName, boolean useStreaming) throws IOException {
        try (InputStream is = Issue81Test.class.getResourceAsStream(inResourceFileName);
            OutputStream os = new FileOutputStream(outFileName)) {
            exportExcel(model, is, os, useStreaming);
        }
    }

    private void exportExcel(Map<String, Object> model, InputStream is, OutputStream os, boolean useStreaming) throws IOException {
        Context context = PoiTransformer.createInitialContext();
        if (model != null) {
            for (String key : model.keySet()) {
                context.putVar(key, model.get(key));
            }
        }
        JexlBuilder jexlBuilder = new JexlBuilder();

        Map<String, Object> funcs = new HashMap<>();
        funcs.put("unit", new JxlsReportModelUtil());
        jexlBuilder.namespaces(funcs);

        JexlEngine jexl = jexlBuilder.create();
        JxlsHelper jxlsHelper = JxlsHelper.getInstance();
        Transformer transformer;
        Workbook wb = null;
        if (useStreaming) {
            wb = WorkbookFactory.create(is);
            context.getConfig().setIsFormulaProcessingRequired(false);

            transformer = new SelectSheetsForStreamingPoiTransformer(wb);
            Set<String> streamingSheets = new HashSet<>();
            streamingSheets.add("Template");
            ((SelectSheetsForStreamingPoiTransformer) transformer).setDataSheetsToUseStreaming(streamingSheets);
            ((PoiTransformer) transformer).setOutputStream(os);
        } else {
            transformer = jxlsHelper.createTransformer(is, os);
        }

        JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig()
                .getExpressionEvaluator();
        evaluator.setJexlEngine(jexl);
        jxlsHelper.processTemplate(context, transformer);

        if (useStreaming) {
            PoiTransformer poiTransformer = (PoiTransformer) transformer;
            wb.setForceFormulaRecalculation(true);
            if (poiTransformer.getWorkbook() instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) poiTransformer.getWorkbook()).dispose();
            }
            wb.close();
        }
    }

    private List<Map<String, Object>> createList() {
        List<Map<String, Object>> list = new ArrayList<>();
        Map<String, Object> item = null;
        List<String> sizList = new ArrayList<>();
        sizList.add("S");
        sizList.add("M");

        item = new LinkedHashMap<>();
        item.put("item", "12345");
        item.put("itemName", "12345-Name");
        item.put("colId", "1231");
        item.put("sizId", "S");
        item.put("qty1", 1);
        item.put("qty2", 2);
        item.put("sizData", sizList);
        list.add(item);

        item = new LinkedHashMap<>();
        item.put("item", "12345");
        item.put("itemName", "12345-Name");
        item.put("colId", "1231");
        item.put("sizId", "M");
        item.put("qty1", 3);
        item.put("qty2", 4);
        item.put("sizData", sizList);
        list.add(item);

        item = new LinkedHashMap<>();
        item.put("item", "12345");
        item.put("itemName", "12345-Name");
        item.put("colId", "1232");
        item.put("sizId", "S");
        item.put("qty1", 5);
        item.put("qty2", 6);
        item.put("sizData", sizList);
        list.add(item);

        item = new LinkedHashMap<>();
        item.put("item", "12345");
        item.put("itemName", "12345-Name");
        item.put("colId", "1232");
        item.put("sizId", "M");
        item.put("qty1", 7);
        item.put("qty2", 8);
        item.put("sizData", sizList);
        list.add(item);

        return list;
    }

    public static class JxlsReportModelUtil {
        
        public WritableCellValue mergeCell(String value, Integer mergerRows) {
            return new MergeCellValue(value, mergerRows);
        }
    }

    public static class MergeCellValue implements WritableCellValue {
        private String value;
        private Integer mergerRows;

        public MergeCellValue(String value, Integer mergerRows) {
            this.value = value;
            this.mergerRows = mergerRows;
        }

        @Override
        public Object writeToCell(Cell cell, Context context) {
            cell.setCellValue(value);
            if (mergerRows == null || mergerRows.intValue() == 0) {
                return cell;
            }
            int rowIndex = cell.getRowIndex();
            Sheet sheet = cell.getSheet();
            int cellIndex = cell.getColumnIndex();
            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, cellIndex, cellIndex + mergerRows - 1));
            return cell;
        }
    }
}
