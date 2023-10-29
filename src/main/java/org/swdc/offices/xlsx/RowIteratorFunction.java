package org.swdc.offices.xlsx;

public interface RowIteratorFunction<E> {

    ExcelCell accept(ExcelCell cell, E element);

}
