package ru.protei.oxmldoc.common.functions;

import java.util.function.Supplier;

public class CustomFunction implements Supplier<String> {
    private final String expression;

    /**
     * @param expression
     */
    public CustomFunction(String expression) {
        this.expression = expression;
    }

    @Override
    public String get() {
        return expression;
    }
}
