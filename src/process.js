import { add_col_aoa, del_col_aoa, parseNum, aoa_ColumnSpliceBySignature, aoa_CopyColumnBySignature, aoa_EnumerateDataBySignature } from "./utils";


export const schema = [
    {
        id: '5.09', title_cells: ["A2"],
        buy: { output: '0000080 5_09 Книга покупок.xls', 'del-rows': [11, 9], 'del-cols': ['O'] },
        sell: { output: '0000090 5_09 Книга продаж.xls', 'del-rows': [11, 9], 'del-cols': ['AU', 'AT', 'AR', 'AP', 'AN', 'AL', 'AJ', 'AH', 'AF', 'AD', 'AB', 'Z', 'X', 'V', 'T', 'R', 'P', 'N', 'L', 'I', 'H', 'F', 'D'] }
    },
    {
        id: '5.10', title_cells: ["A2", "G1", "H1"],
        buy: {
            output: '0000080 5_10 Книга покупок.xls', 
            // PREFERABLE "process" workflow: process, splice columns, splice rows
            process: (aoa) => 
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
        sell: { output: '0000090 5_10 Книга продаж.xls', 
        // PREFERABLE "process" workflow: process, splice columns, splice rows
            process: (aoa) => {
                const DATA_IS_PRESENT_COL = 0

                // меняем формат ИИН и КПП
                const IIN_COL = 6, KPP_COL = 7; 
                const DATA_STARTS_ROW = 11; 
                for (let r = DATA_STARTS_ROW; r < aoa.length; r++) {
                    if (aoa[r][DATA_IS_PRESENT_COL] != "") {
                        aoa[r][IIN_COL] = aoa[r][IIN_COL] + "/" + aoa[r][KPP_COL]
                    }
                }
                const ENUM_ROW = 10;
                aoa_ColumnSpliceBySignature(aoa, ENUM_ROW, "3а", 1, 4);     // удалить 3а и добавить 4 колонки на её месте
                aoa_ColumnSpliceBySignature(aoa, ENUM_ROW, "5а", 0, 7);     // вставить 7 колонок перед пунктом 5а
                aoa_CopyColumnBySignature(aoa, ENUM_ROW, "8а", "row-13", DATA_STARTS_ROW); // расставляем колонки по местам
                aoa_CopyColumnBySignature(aoa, ENUM_ROW, "5а", "row-14", DATA_STARTS_ROW);
                aoa_CopyColumnBySignature(aoa, ENUM_ROW, "6а", "row-15", DATA_STARTS_ROW);
                aoa_CopyColumnBySignature(aoa, ENUM_ROW, "7",  "row-16", DATA_STARTS_ROW);
                aoa_CopyColumnBySignature(aoa, ENUM_ROW, "8б", "row-17", DATA_STARTS_ROW);
                aoa_CopyColumnBySignature(aoa, ENUM_ROW, "5б", "row-18", DATA_STARTS_ROW);
                aoa_CopyColumnBySignature(aoa, ENUM_ROW, "6б", "row-19", DATA_STARTS_ROW);
                aoa_ColumnSpliceBySignature(aoa, ENUM_ROW, "5а", 0, 5);     // добавить ещё 5 колонок
                aoa_ColumnSpliceBySignature(aoa, ENUM_ROW, "1", 0, 2, "Продажа", "Руководитель");  // сдвинуть всё между "Продажа" и "Руководитель" на две колонки
                aoa_EnumerateDataBySignature(aoa, ENUM_ROW, "row-0", "1", DATA_STARTS_ROW);  // добавить нумерацию строк
                const OUTPUT_COL_NUM = 27
                for (let i = 0; i < OUTPUT_COL_NUM; i++) {  // добавить нумерацию столбцов
                    aoa[ENUM_ROW][i] = i+1;    
                }
                aoa.splice(2, 0, [""]);     // добавить пустую сторку перед третьей
                aoa.splice(8, 2)            // удалить две строки после 8-й
            } 
        }
    }
];