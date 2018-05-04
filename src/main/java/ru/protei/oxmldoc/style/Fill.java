package ru.protei.oxmldoc.style;

import ru.protei.oxmldoc.style.enums.FillPattern;

public class Fill {
    private FillPattern pattern = FillPattern.NONE;
    private Color backgroundColor;
    private Color frontColor;
    private final int id;

    public Fill (int id){
        this.id = id;
    }

    /**
     * @return id
     */
    public int getId()
    {
        return id;
    }

    public FillPattern getPattern() {
        return pattern;
    }

    public void setPattern(FillPattern pattern) {
        this.pattern = pattern;
    }

    public Color getBackgroundColor() {
        return backgroundColor;
    }

    public void setBackgroundColor(Color backgroundColor) {
        this.backgroundColor = backgroundColor;
    }

    public Color getFrontColor() {
        return frontColor;
    }

    public void setFrontColor(Color frontColor) {
        this.frontColor = frontColor;
    }
}
