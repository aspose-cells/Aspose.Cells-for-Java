package com.aspose.gridjsdemo.usermanagement.exception;

 
public class CustomFieldValidationException extends Exception {
    private static final long serialVersionUID = 1L;

    private String fieldName;

    public CustomFieldValidationException(String message, String fieldName) {
        super(message);
        this.fieldName = fieldName;
    }

    public String getFieldName() {
        return this.fieldName;
    }
}