package ru.protei.oxmldoc.common

enum class RowsLimitRule {

    /**
     * Ignore limit value, allow to append any number of rows
     */
    IGNORE,

    /**
     * Skip all rows above limit
     */
    SKIP,

    /**
     * Generate error
     */
    ERROR,

    /**
     * Automatically create a new worksheet and proceed append to it
     */
    SPLIT
}