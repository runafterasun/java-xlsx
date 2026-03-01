package ru.objectsfill.exception;

/**
 * Exception thrown when an Excel import operation fails.
 * Wraps lower-level I/O, reflection, or parsing errors with a meaningful context message.
 */
public class ExcelImportException extends RuntimeException {

    /**
     * Constructs an exception with the given message.
     *
     * @param message description of the failure
     */
    public ExcelImportException(String message) {
        super(message);
    }

    /**
     * Constructs an exception with the given message and underlying cause.
     *
     * @param message description of the failure
     * @param cause   the underlying exception
     */
    public ExcelImportException(String message, Throwable cause) {
        super(message, cause);
    }
}
