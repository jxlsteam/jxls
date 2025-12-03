# Link

Use this command for render a hyperlink.

```
jx:link(href="https://jxls.sourceforge.net" label="Jxls" lastCell="A2")
```

`href`: target address. The target address can be an URL or an Excel cellref.

`label`: visible text of the hyperlink

`type`: default value is "URL". Set to "DOCUMENT" if you want to point to an Excel cell. It must be a name of the org.apache.poi.common.usermodel.HyperlinkType enum.

`color`: color of the visible text, default is BLUE. It must be a name of the org.apache.poi.ss.usermodel.IndexedColors enum.

`lastCell`: area end (same cell)

The attributes href, label, type and color can be expressions or values.

jx:link is part of jxls-poi and is only available if you use JxlsPoiTemplateFillerBuilder or add the LinkCommand using withCommand().

This command is a community contribution.
