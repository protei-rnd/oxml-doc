package ru.protei.oxmldoc.style;

import ru.protei.oxmldoc.style.enums.HorizontalAligment;
import ru.protei.oxmldoc.style.enums.VerticalAligment;

public class Aligment
{
    private HorizontalAligment horizAlign;
    private Boolean wrapText;
    private VerticalAligment vertAlign;

    /**
     * @param horizAlign
     * @param wrapText
     * @param vertAlign
     */
    public Aligment(HorizontalAligment horizAlign, Boolean wrapText, VerticalAligment vertAlign)
    {
        this.horizAlign = horizAlign;
        this.wrapText = wrapText;
        this.vertAlign = vertAlign;
    }

    /**
     * @param horizAlign
     */
    public Aligment(HorizontalAligment horizAlign)
    {
        this.horizAlign = horizAlign;
        this.wrapText = false;
        this.vertAlign = VerticalAligment.TOP;
    }

    public HorizontalAligment getHorizAlign() {
        return horizAlign;
    }

    public void setHorizAlign(HorizontalAligment horizAlign) {
        this.horizAlign = horizAlign;
    }

    public Boolean getWrapText() {
        return wrapText;
    }

    public void setWrapText(Boolean wrapText) {
        this.wrapText = wrapText;
    }

    public VerticalAligment getVertAlign() {
        return vertAlign;
    }

    public void setVertAlign(VerticalAligment vertAlign) {
        this.vertAlign = vertAlign;
    }
}
