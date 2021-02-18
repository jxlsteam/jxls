package org.jxls.extractors;

import java.util.ArrayList;
import java.util.List;
import org.junit.Test;
import static org.junit.Assert.assertEquals;
import org.jxls.extractors.LiteralsExtractor;

/**
 * @author Alexander Lust
 */

public class LiteralsExtractorTest {

    /**
     * Tests whether the LiteralsExtractor is used correctly.
     */	
	@Test
    public void testExtract() {
    	List<String> literalList = new ArrayList<String>();
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
    	
    	String text = "Autor:\r\n"
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
        LiteralsExtractor extractor = new LiteralsExtractor(text, literalList);
        literalList = extractor.extract();
        
        assertEquals("Number of literalList is wrong", 2, literalList.size());
        assertEquals("First member of literalList is wrong", "Autor:\r", literalList.get(0));
        assertEquals("Second member of literalList is wrong", expectedText, literalList.get(1));
    }
	
    /**
     * Tests null.
     */	
	@Test(expected = NullPointerException.class)
    public void testNullPointerException() {
    	List<String> literalList = null;
    	String text = null;
        LiteralsExtractor extractor = new LiteralsExtractor(text, literalList);
        literalList = extractor.extract();
    }
    
}
