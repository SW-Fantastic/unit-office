package org.swdc.offices;

import org.swdc.offices.xlsx.ExcelCell;

public interface RowIteratorFunction<E,C> {

    C accept(C cell, E element);

}
