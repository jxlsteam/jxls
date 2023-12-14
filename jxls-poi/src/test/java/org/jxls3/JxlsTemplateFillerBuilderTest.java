package org.jxls3;

import java.io.InputStream;

import org.junit.Test;
import org.jxls.builder.JxlsTemplateFillerBuilder;
import org.jxls.common.JxlsException;
import org.jxls.transform.poi.PoiTransformerFactory;
import org.jxls.util.CannotOpenWorkbookException;

public class JxlsTemplateFillerBuilderTest {

	@Test(expected = IllegalArgumentException.class)
	public void areaBuilderIsNull() {
		JxlsTemplateFillerBuilder.newInstance().withAreaBuilder(null);
	}

	@Test(expected = IllegalArgumentException.class)
	public void expressionEvaluatorFactoryIsNull() {
		JxlsTemplateFillerBuilder.newInstance().withExpressionEvaluatorFactory(null);
	}

	@Test(expected = IllegalArgumentException.class)
	public void transformerFactoryIsNull() {
		JxlsTemplateFillerBuilder.newInstance().withTransformerFactory(null);
	}

	@Test(expected = CannotOpenWorkbookException.class)
	public void inputStreamIsNull() {
		JxlsTemplateFillerBuilder.newInstance().withTemplate((InputStream) null);
	}

	@Test(expected = JxlsException.class)
	public void withTransformerFactoryNotCalled() {
		JxlsTemplateFillerBuilder.newInstance().build();
	}

	@Test(expected = JxlsException.class)
	public void withTemplateNotCalled() {
		JxlsTemplateFillerBuilder.newInstance().withTransformerFactory(new PoiTransformerFactory()).build();
	}
}
