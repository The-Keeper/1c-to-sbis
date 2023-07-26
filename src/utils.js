
/**
 * Deletes a column from array-of-arrays.
 *
 * @param {Array<any>} aoa - Array to delete the column from
 * @param {number} col_num - zero-based index of column to delete
 *
 * @example
 *
 *     del_col_aoa(aoa, 3)
 */
export function del_col_aoa(aoa, col_num) {
    for (let i in aoa) {
        aoa[i].splice(col_num, 1);
    }
}

/**
 * Deletes a row from array-of-arrays.
 *
 * @param {Array<any>} aoa - Array to delete the row from
 * @param {number} row_num - zero-based index of row to delete
 *
 * @example
 *
 *     del_row_aoa(aoa, 3)
 */
export function del_row_aoa(aoa, row_num) {
    aoa.splice(row_num, 1);
}


/**
 * Adds a column to array-of-arrays.
 *
 * @param {Array<any>} aoa - Array to add the column to
 * @param {number} col_num - zero-based index of column to add after
 *
 * @example
 *
 *     add_col_aoa(aoa, 3)
 */
export function add_col_aoa(aoa, col_num) {
    for (let row in aoa) {
        aoa[row].splice(col_num, 0, "")
    }      
}


/**
 * convert comma-separated number string to a number
 * @param {string} s 
 */
export function parseNum(s) {
    s = s.toString();
    if (!s) {
        return 0;
    }
    let result = parseFloat(s.replace(',', '.'))
    if (result) {
        return result;
    }
    return 0;
}

/**
 * Enumerates rows into given column by provided signatures
 * @param {Array<any>} aoa - array-of-arrays to perform the action on
 * @param {number} row - row to look for signatures in
 * @param {string} enumerationColSignature - signature of column where to put enumeration
 * @param {string} dataColSignature - signature of column with contiguous chunk of data
 * @param {number} rowDataStart - row where data starts
 */
export function aoa_EnumerateDataBySignature(aoa, row, enumerationColSignature, dataColSignature, rowDataStart) {
    let col_data = -1, col_to = -1; 
    for (let c = 0; c < aoa[row].length; c++) {
        if (aoa[row][c] == dataColSignature) 
            col_data = c;
        if (aoa[row][c] == enumerationColSignature) 
            col_to = c;
    }
    if (col_data>=0 && col_to>=0 && col_data !=col_to) {
        for (let i=0; rowDataStart + i < aoa.length; i++) {
            let r = rowDataStart + i;
            if (aoa[r][col_data] == "") 
                break;
            aoa[r][col_to] = i+1;
            }
        }
}


/**
 * Copies data between columns based on signature
 * @param {Array<any>} aoa - array-of-arrays to perform the action on
 * @param {number} row - row to look for signatures in
 * @param {string} fromColSignature - signature of column where to get data from
 * @param {string} toColSignature - signature of column where to put data data
 * @param {number} rowDataStart - row where data starts
 */
export function aoa_CopyColumnBySignature(aoa, row, fromColSignature, toColSignature, rowDataStart) {
    let col_from = -1, col_to = -1; 
    for (let c = 0; c < aoa[row].length; c++) {
        if (aoa[row][c] == fromColSignature) 
            col_from = c;
        if (aoa[row][c] == toColSignature) 
            col_to = c;
        }
    if (col_from>=0 && col_to>=0 && col_from !=col_to) {
        for (let r = rowDataStart; r < aoa.length; r++) {
            if (aoa[r][0] == "") 
                break;
            aoa[r][col_to] = aoa[r][col_from];
            }
        }
}

/**
 * Splices the columns of array-of-arrays object based on value of a cell in a given row
 * @param {Array<any>} aoa - array-of-arrays to perform the action on
 * @param {number}  row - row to scan left-to-right
 * @param {string} signature - string that the cell must contain to perform operations on
 * @param {number} delCount - number of columns to delete
 * @param {number} addCount - number of columns to insert
 * @param {string} [startAfterSignature=""] - start moving columns after row that starts with given signature
 * @param {string} [terminateBeforeSignature=""] - finish moving columns before row that ends with given signature
 */
export function aoa_ColumnSpliceBySignature (aoa, row, signature, delCount, addCount, startAfterSignature = "", terminateBeforeSignature = "") {   
    for (let c = 0; c < aoa[row].length; c++) {
        const partialMoveAllowed = startAfterSignature && terminateBeforeSignature && !(delCount>0);
        if (aoa[row][c] == signature) {
            const spliceArgs = [c, delCount].concat(Array(addCount).fill(""));
            const additionArgs = [aoa[row].length, 0].concat(Array(addCount).fill(""));
            const signAddArgs  = [c, delCount].concat([ ...Array(addCount).keys() ].map(  i => `row-${i+c}` ))

            let partialMoveStarted = false;
            for (let r=0; r < aoa.length; r++) {

                if (partialMoveAllowed && aoa[r][0].startsWith(terminateBeforeSignature)) {
                    partialMoveStarted = false;
                }
                if (partialMoveAllowed && !partialMoveStarted) {
                    aoa[r].splice(...additionArgs);
                } else                 
                {
                    if (r == row) {
                        aoa[r].splice(...signAddArgs);                     
                    } else {
                        aoa[r].splice(...spliceArgs);
                    }
                } 
                if (partialMoveAllowed && aoa[r][0].startsWith(startAfterSignature)) {
                    partialMoveStarted = true;
                }
            }
            break;
        }
    }
}