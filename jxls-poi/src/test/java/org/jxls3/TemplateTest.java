package org.jxls3;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.io.IOUtils;
import org.junit.Test;
import org.jxls.TestWorkbook;
import org.jxls.common.JxlsException;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class TemplateTest {
	private File outputFile;
	private File templateFile;
	
	@Test
	public void url() throws IOException {
		outputFile("url");
		URL url = EachTest.class.getResource(EachTest.class.getSimpleName() + ".xlsx");
		JxlsPoiTemplateFillerBuilder.newInstance().withTemplate(url).buildAndFill(data(), outputFile);
        verify();
	}
	
	@Test(expected = IllegalArgumentException.class)
	public void urlNull() throws IOException {
		JxlsPoiTemplateFillerBuilder.newInstance().withTemplate((URL) null);
	}

	@Test
	public void file() throws IOException {
		outputFile("file");
        templateFile();
		try {
			JxlsPoiTemplateFillerBuilder.newInstance().withTemplate(templateFile).buildAndFill(data(), outputFile);
	        verify();
		} finally {
			templateFile.delete();
		}
	}

	@Test(expected = IllegalArgumentException.class)
	public void fileNull() throws FileNotFoundException {
		JxlsPoiTemplateFillerBuilder.newInstance().withTemplate((File) null);
	}

	@Test(expected = JxlsException.class)
	public void isNotFile() throws FileNotFoundException {
		File file = new File("target");
		file.mkdir();
		JxlsPoiTemplateFillerBuilder.newInstance().withTemplate(file);
	}

	@Test
	public void filename() throws IOException {
		outputFile("filename");
        templateFile();
		try {
			JxlsPoiTemplateFillerBuilder.newInstance().withTemplate(templateFile.getAbsolutePath()).buildAndFill(data(), outputFile);
	        verify();
		} finally {
			templateFile.delete();
		}
	}

	@Test(expected = IllegalArgumentException.class)
	public void filenameNull() throws FileNotFoundException {
		JxlsPoiTemplateFillerBuilder.newInstance().withTemplate((String) null);
	}

	@Test(expected = IllegalArgumentException.class)
	public void filenameBlank() throws FileNotFoundException {
		JxlsPoiTemplateFillerBuilder.newInstance().withTemplate(" \t");
	}

	private void outputFile(String name) {
        outputFile = new File("target/TemplateTest_" + name + "_output.xlsx");
        outputFile.getParentFile().mkdir();
	}

	private void templateFile() throws IOException, FileNotFoundException {
		templateFile = new File("target/TemplateTest_template.xlsx");
		IOUtils.copy(EachTest.class.getResourceAsStream(EachTest.class.getSimpleName() + ".xlsx"),
				new FileOutputStream(templateFile));
	}

	private Map<String, Object> data() {
		Map<String, Object> data = new HashMap<>();
	    data.put("employees", Employee.generateSampleEmployeeData());
		return data;
	}

	private void verify() {
		// Verify
		try (TestWorkbook w = new TestWorkbook(outputFile)) {
		    w.selectSheet("Employees");
		    assertEquals("Elsa", w.getCellValueAsString(2, 1)); 
		}
	}
}
