package org.jxls.templatebasedtests;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

/**
 * Github issue 17: Image and multi sheet
 */
public class Issue17Test {

    @Test
    public void test() throws Exception {
        // Prepare
        Context context = new Context();
        context.putVar("sheets", getItems());
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        // with evaluate formulas = true
        tester.processTemplate(context);
        
        // Verify
        // Each sheet with an image?
    }

    private List<Item> getItems() {
        List<Item> items = new ArrayList<>();
        Item item = new Item();
        item.setId("registered_user");
        item.setCountry_code("JP");
        item.setApp_type("ANDROID");
        items.add(item);
        item = new Item();
        item.setId("registered_user1");
        item.setCountry_code("KO");
        item.setApp_type("ANDROID");
        items.add(item);
        item = new Item();
        item.setId("registered_user2");
        item.setCountry_code("US");
        item.setApp_type("ANDROID");
        items.add(item);
        return items;
    }

    public static class Item {
        private String id;
        private String country_code;
        private String app_type;

        public String getId() {
            return id;
        }

        public void setId(String id) {
            this.id = id;
        }

        public String getCountry_code() {
            return country_code;
        }

        public void setCountry_code(String country_code) {
            this.country_code = country_code;
        }

        public String getApp_type() {
            return app_type;
        }

        public void setApp_type(String app_type) {
            this.app_type = app_type;
        }
    }    
}
