package org.jxls.transform.poi;

import org.jxls.builder.JxlsTemplateFillerBuilder;
import org.jxls.command.ImageCommand;
import org.jxls.command.MergeCellsCommand;
import org.jxls.common.PoiExceptionLogger;
import org.jxls.common.PoiExceptionThrower;

public class JxlsPoiTemplateFillerBuilder extends JxlsTemplateFillerBuilder<JxlsPoiTemplateFillerBuilder> {

    public JxlsPoiTemplateFillerBuilder() {
        withLogger(new PoiExceptionLogger());
        withTransformerFactory(new PoiTransformerFactory());
        withCommand(ImageCommand.COMMAND_NAME, ImageCommand.class);
        withCommand(MergeCellsCommand.COMMAND_NAME, MergeCellsCommand.class);
    }
    
    public static JxlsPoiTemplateFillerBuilder newInstance() {
        return new JxlsPoiTemplateFillerBuilder();
    }
    
    public JxlsPoiTemplateFillerBuilder withExceptionThrower() {
    	withLogger(new PoiExceptionThrower());
    	return this;
    }
}
