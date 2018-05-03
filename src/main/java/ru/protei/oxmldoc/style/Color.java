package ru.protei.oxmldoc.style;

public class Color {
    private final int rgb;

    public Color (int rgb){
        this.rgb = rgb;
    }

    public Color (java.awt.Color awtColor){
        rgb = awtColor.getRGB();
    }

    public int getRgb() {
        return rgb;
    }
}
