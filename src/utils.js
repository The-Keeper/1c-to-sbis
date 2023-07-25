
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
