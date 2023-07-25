import { add_col_aoa, del_col_aoa, parseNum } from "./utils";

export const schema = [
    {
        id: '5.09', title_cells: ["A2"],
        buy: { output: '0000080 5_09 Книга покупок.xls', 'del-rows': [11, 9], 'del-cols': ['O'] },
        sell: { output: '0000090 5_09 Книга продаж.xls', 'del-rows': [11, 9], 'del-cols': ['AU', 'AT', 'AR', 'AP', 'AN', 'AL', 'AJ', 'AH', 'AF', 'AD', 'AB', 'Z', 'X', 'V', 'T', 'R', 'P', 'N', 'L', 'I', 'H', 'F', 'D'] }
    },
    {
        id: '5.10', title_cells: ["A2", "G1", "H1"],
        buy: {
            output: '0000080 5_10 Книга покупок.xls', process: (aoa) => 
            {
                aoa.splice(1, 0, [""]);     // insert one empty row
                aoa.splice(8, 2)            // delete two rows after 8th
                
                add_col_aoa(aoa, 1)         // add column

                // cols 9/10 = ИИН/КПП
                // row 10     - first significant row
                for (let r = 10; r < aoa.length; r++) {
                    if (aoa[r][0] != "") {
                        aoa[r][9] = aoa[r][9] + "/" + aoa[r][10]

                        // НДС
                        aoa[r][14] = parseNum(aoa[r][15]) + parseNum(aoa[r][17]) + parseNum(aoa[r][19])
                    }
                }
                del_col_aoa(aoa, 10); // delete the no longer needed КПП column
                del_col_aoa(aoa, 12);
                add_col_aoa(aoa, 11); // add two colums after
                add_col_aoa(aoa, 11);
                for (let i=4; i>0; i--) {   // add four columns
                    add_col_aoa(aoa, 15);
                }

                for (let i = 0; i < 19; i++) {
                    aoa[9][i] = i+1     // add numbers to enumeration row
                }
            }
        },
        sell: { output: '0000090 5_10 Книга продаж.xls', }
    }
];