package org.jxls.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.w3c.dom.Text;
import org.xml.sax.SAXException;

/**
 * With R{key} in the Excel XLSX template file, resource bundles can be accessed to realize multilingualism.
 * This is the JXLS standard solution since 2.8.0. It is particularly necessary when using PivotTables.
 * Usually this class is called before the actual JXLS processing in order to create
 * a modified temporary template file from the given template file.
 * With R{key=defaultValue}, a default value can be specified for the case that there is no value for the key.
 * Otherwise key is the fallback value.
 * The notation can be adjusted using setStart, setEnd and setDefaultValueDelimiter.
 */
public abstract class JxlsNationalLanguageSupport {
    private String start = "R{";
    private String end = "}";
    private String defaultValueDelimiter = "=";
    private Pattern pattern;
    
    /**
     * @param in XLSX input stream
     * @return new temp file. Caller must delete it.
     */
    public File process(InputStream in) {
        try {
            File out = File.createTempFile("JXLS-R-", ".xlsx");
            process(in, new FileOutputStream(out));
            return out;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * @param in XLSX input stream that contain R{key} elements
     * @param out XLSX output stream for writing the result that contains translated R{key} elements
     * @throws IOException -
     * @throws TransformerConfigurationException -
     * @throws ParserConfigurationException -
     * @throws SAXException -
     * @throws TransformerException -
     * @throws TransformerFactoryConfigurationError -
     */
    public void process(InputStream in, OutputStream out) throws IOException, TransformerConfigurationException, ParserConfigurationException, SAXException, TransformerException, TransformerFactoryConfigurationError {
        pattern = Pattern.compile(Pattern.quote(start) + "(.*?)" + Pattern.quote(end));
        try (ZipInputStream zipin = new ZipInputStream(in); ZipOutputStream zipout = new ZipOutputStream(out)) {
            ZipEntry zipEntry;
            while ((zipEntry = zipin.getNextEntry()) != null) {
                processZipEntry(zipEntry, zipin, zipout);
            }
        }
        pattern = null;
    }

    protected void processZipEntry(ZipEntry zipEntry, InputStream in, ZipOutputStream zipout) throws IOException, ParserConfigurationException, SAXException, TransformerException {
        zipout.putNextEntry(new ZipEntry(zipEntry.getName()));
        if (zipEntry.getName().toLowerCase().endsWith(".xml")) {
            DocumentBuilder builder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
            Document dom = builder.parse(getNoCloseInputStream(in));
            processElement(dom.getDocumentElement());
            javax.xml.transform.TransformerFactory.newInstance().newTransformer().transform(new DOMSource(dom), new StreamResult(zipout));
        } else {
            transfer(in, zipout);
        }
        zipout.closeEntry();
    }

    private InputStream getNoCloseInputStream(final InputStream in) {
        return new InputStream() {
            @Override
            public int read() throws IOException {
                return in.read();
            }
            
            @Override
            public void close() throws IOException { // Do not close after reading a single Zip file entry.
            }
        };
    }
    
    private void processElement(Element root) {
        // Attributes
        NamedNodeMap attributes = root.getAttributes();
        for (int i = 0; i < attributes.getLength(); i++) {
            Node item = attributes.item(i);
            if (item instanceof Attr) {
                String val = ((Attr) item).getValue();
                String newValue = translateAll(val);
                if (!val.equals(newValue)) {
                    ((Attr) item).setValue(newValue);
                }
            }
        }

        // Text and sub elements
        NodeList children = root.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node item = children.item(i);
            if (item instanceof Text) {
                String val = ((Text) item).getTextContent();
                String newValue = translateAll(val);
                if (!val.equals(newValue)) {
                    item.setTextContent(newValue);
                }
            } else if (item instanceof Element) {
                processElement((Element) item); // recursive
            }
        }
    }

    protected String translateAll(String text) {
        StringBuffer ret = new StringBuffer();
        Matcher matcher = pattern.matcher(text);
        while (matcher.find()) {
            String name = matcher.group(1);
            String fallback = name;
            int o = name.indexOf(defaultValueDelimiter);
            if (o > 0) {
                fallback = name.substring(o + defaultValueDelimiter.length());
                name = name.substring(0, o);
            }
            String newValue = translate(name.trim(), fallback);
            matcher.appendReplacement(ret, newValue);
        }
        matcher.appendTail(ret);
        return ret.toString();
    }

    protected abstract String translate(String name, String fallback);

    protected void transfer(InputStream in, OutputStream out) throws IOException {
        byte[] buf = new byte[8192];
        int len;
        while ((len = in.read(buf)) > 0) {
            out.write(buf, 0, len);
        }
    }
    
    public String getStart() {
        return start;
    }

    public void setStart(String start) {
        this.start = start;
    }

    public String getEnd() {
        return end;
    }

    public void setEnd(String end) {
        this.end = end;
    }

    public String getDefaultValueDelimiter() {
        return defaultValueDelimiter;
    }

    public void setDefaultValueDelimiter(String defaultValueDelimiter) {
        this.defaultValueDelimiter = defaultValueDelimiter;
    }
}
