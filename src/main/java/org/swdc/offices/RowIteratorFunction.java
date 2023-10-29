package org.swdc.offices;

import org.swdc.offices.xlsx.ExcelCell;

public interface RowIteratorFunction<E> {

    ExcelCell accept(ExcelCell cell, E element);

}
