package org.wuwz.poi.pojo;

public class ExportItem
{
    private String field;
    private String display;
    private short width;
    private String convert;
    private short color;
    private String replace;
    private String range;

    private String dataType;

    public String getField()
    {
        return this.field;
    }

    public ExportItem setField(String field)
    {
        this.field = field;
        return this;
    }

    public String getDisplay()
    {
        return this.display;
    }

    public ExportItem setDisplay(String display)
    {
        this.display = display;
        return this;
    }

    public short getWidth()
    {
        return this.width;
    }

    public ExportItem setWidth(short width)
    {
        this.width = width;
        return this;
    }

    public String getConvert()
    {
        return this.convert;
    }

    public ExportItem setConvert(String convert)
    {
        this.convert = convert;
        return this;
    }

    public short getColor()
    {
        return this.color;
    }

    public ExportItem setColor(short color)
    {
        this.color = color;
        return this;
    }

    public String getReplace()
    {
        return this.replace;
    }

    public ExportItem setReplace(String replace)
    {
        this.replace = replace;
        return this;
    }

    public String getRange()
    {
        return this.range;
    }

    public ExportItem setRange(String range)
    {
        this.range = range;
        return this;
    }

    public String getDataType() {
        return dataType;
    }

    public ExportItem setDataType(String dataType) {
        this.dataType = dataType;
        return this;
    }
}
