package ru.protei.oxmldoc.reader;

public interface ExcelHandler {
    /**
     * This method is called when new worksheet is processed
     *
     * @param name worksheet name
     */
    public void beginWorksheet(String name);

    /**
     * This method is called when full worksheet was processed
     */
    public void endWorksheet();

    /**
     * This method is called when each new row is processed
     *
     * @param rowNumber row number (beginning with 0)
     * @param min       the minimum column index (beginning with 0)
     * @param max       the maximum column index (beginning with 0)
     */
    public void beginRow(int rowNumber, int min, int max);

    /**
     * This method is called when full row was processed
     */
    public void endRow();

    /**
     * This method is called when a value is assigned to cell in current row
     *
     * @param cell  cell number (beginning with 0)
     * @param value cell value, can be NULL if cell is empty
     */
    public void setCell(int cell, CellValue value);
}
