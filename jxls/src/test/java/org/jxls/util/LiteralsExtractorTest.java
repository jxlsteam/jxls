package org.jxls.util;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.util.List;

import org.junit.Test;

/**
 * @author Alexander Lust
 */
public class LiteralsExtractorTest {

    @Test
    public void testSimple() {
        String expectedText = "jx:each(items=\"jdbc.query('select * from RS_USER')\" var=\"user\" lastCell=\"D4\")";
        String text = "author:\r\n"
                + "jx:each(items=\"jdbc.query('select * from RS_USER')\" var=\"user\" lastCell=\"D4\")";

        List<String> literalList = new LiteralsExtractor().extract(text);
        
        assertEquals("Number of literalList is wrong", 2, literalList.size());
        assertEquals("First member of literalList is wrong", "author:\r", literalList.get(0));
        assertEquals("Second member of literalList is wrong", expectedText, literalList.get(1));
    }
    
    /** This especially tests a leading space. */
    @Test
    public void testWhitespace() {
        String expectedText = "jx:each(items=\"jdbc.query('select * from RS_USER')\" var=\"user\" lastCell=\"D4\")";
        String text = "author:\r\n"
                + " \t " // leading whitespace
                + "jx:each(items=\"jdbc.query('select * from RS_USER')\" var=\"user\" lastCell=\"D4\")"
                + " \t \n"; // trailing whitespace

        List<String> literalList = new LiteralsExtractor().extract(text);
        
        assertEquals("Number of literalList is wrong", 2, literalList.size());
        assertEquals("First member of literalList is wrong", "author:\r", literalList.get(0));
        assertEquals("Second member of literalList is wrong", expectedText, literalList.get(1));
    }

    @Test
    public void testTwoCommands() {
        String text = "author:\r\n"
                + "jx:each(items=\"list1\" var=\"i\" lastCell=\"D4\")\r\n"
                + "jx:each(items=\"jdbc.query('select * from RS_USER')\" var=\"user\" lastCell=\"D4\")\r\n";

        List<String> literalList = new LiteralsExtractor().extract(text);
        
        assertEquals("Number of literalList is wrong", 3, literalList.size());
        assertEquals("jx:each(items=\"list1\" var=\"i\" lastCell=\"D4\")", literalList.get(1));
        assertEquals("jx:each(items=\"jdbc.query('select * from RS_USER')\" var=\"user\" lastCell=\"D4\")", literalList.get(2));
    }

    /**
     * Tests whether the LiteralsExtractor is used correctly.
     */
    @Test
    public void testMultiLineSQL() {
        String expectedText = "jx:each(items=\"jdbc.query('select * \r\n"
                + "-- comment 1\r\n"
                + "\r\n"
                + "/*\r\n"
                + "coment2\r\n"
                + "\r\n"
                + "comment3\r\n"
                + "\r\n"
                + "*/\r\n"
                + "from RS_USER')\" var=\"user\" lastCell=\"D4\")";
        String text = "author:\r\n"
                + "jx:each(items=\"jdbc.query('select * \r\n"
                + "-- comment 1\r\n"
                + "\r\n"
                + "/*\r\n"
                + "coment2\r\n"
                + "\r\n"
                + "comment3\r\n"
                + "\r\n"
                + "*/\r\n"
                + "from RS_USER')\" var=\"user\" lastCell=\"D4\")";
        
        List<String> literalList = new LiteralsExtractor().extract(text);
        
        assertEquals("Number of literalList is wrong", 2, literalList.size());
        assertEquals("First member of literalList is wrong", "author:\r", literalList.get(0));
        assertEquals("Second member of literalList is wrong", expectedText, literalList.get(1));
    }

    @Test(expected = NullPointerException.class)
    public void testNull() {
        new LiteralsExtractor().extract(null);
    }

    @Test
    public void testEmpty() {
        List<String> literalList = new LiteralsExtractor().extract("");

        assertTrue(literalList.isEmpty());
    }

    // issue 111
    @Test
    public void testCommentIf() {
        String expectedText = "jx:each(items=\"jdbc.query('select * from RS_USER')\" var=\"user\" lastCell=\"D4\")";
        String text = "author:\r\n"
                + "jx:each(items=\"jdbc.query('select * from RS_USER')\" var=\"user\" lastCell=\"D4\")\r\n"
                + "//jx:if(condition=\"user.XYZ != 4\" lastCell=\"D4\")";

        List<String> literalList = new LiteralsExtractor().extract(text);
        
        assertEquals("Number of literalList is wrong", 3, literalList.size());
        assertEquals("First member of literalList is wrong", "author:\r", literalList.get(0));
        assertEquals("Second member of literalList is wrong", expectedText, literalList.get(1));
        assertEquals("Third member of literalList is wrong", "//jx:if(condition=\"user.XYZ != 4\" lastCell=\"D4\")", literalList.get(2));
    }
}
