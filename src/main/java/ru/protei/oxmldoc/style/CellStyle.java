package ru.protei.oxmldoc.style;

public class CellStyle
{
    private final int id;

    private Aligment aligment;

    private Font font;

    private Fill fill;

    private NumberFormat numberFormat;

    public CellStyle (int id){
        this.id = id;
    }

    /**
     * @return the id
     */
    public int getId()
    {
        return id;
    }

    public Aligment getAligment() {
        return aligment;
    }

    public void setAligment(Aligment aligment) {
        this.aligment = aligment;
    }

    public Font getFont() {
        return font;
    }

    public void setFont(Font font) {
        this.font = font;
    }

    public Fill getFill() {
        return fill;
    }

    public void setFill(Fill fill) {
        this.fill = fill;
    }

    public NumberFormat getNumberFormat() {
        return numberFormat;
    }

    public void setNumberFormat(NumberFormat numberFormat) {
        this.numberFormat = numberFormat;
    }
}
