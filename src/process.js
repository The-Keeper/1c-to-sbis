import { parseNum, aoa_ColumnSpliceBySignature, aoa_CopyColumnBySignature, aoa_EnumerateDataBySignature } from "./utils";

export const schema = [
    {
        id: '5.09', title_cells: ["A2"],
        buy: { output: '0000080 5_09 Книга покупок.xls', 'del-rows': [11, 9], 'del-cols': ['O'] },
        sell: { output: '0000090 5_09 Книга продаж.xls', 'del-rows': [11, 9], 'del-cols': ['AU', 'AT', 'AR', 'AP', 'AN', 'AL', 'AJ', 'AH', 'AF', 'AD', 'AB', 'Z', 'X', 'V', 'T', 'R', 'P', 'N', 'L', 'I', 'H', 'F', 'D'] }
    },
	{
		id: '5.10',
		title_cells: ['A2', 'G1', 'H1'],
		buy: {
			output: '0000080 5_10 Книга покупок.xls',
			// PREFERABLE "process" workflow: process, splice columns, splice rows
			process: (/** @type {Array<any>} - array-of-arrays extracted from the file  */ aoa) => {
				const DATA_IS_PRESENT_COL = 0;

				// меняем формат ИИН и КПП
				const IIN_COL = 8,
					KPP_COL = 9;
				const DATA_STARTS_ROW = 11;
				const NDS_PART_COLS = [13, 15, 18];
				const NDS_TARGET_TOL = 12;
				for (let r = DATA_STARTS_ROW; r < aoa.length; r++) {
					if (aoa[r][DATA_IS_PRESENT_COL] != '') {
						aoa[r][IIN_COL] = aoa[r][IIN_COL] + '/' + aoa[r][KPP_COL];
						aoa[r][NDS_TARGET_TOL] = NDS_PART_COLS.map((i) => parseNum(aoa[r][i])).reduce(
							(partialSum, a) => partialSum + a,
							0
						);
					}
				}
				const ENUM_ROW = 10,
					OUTPUT_COL_NUM = 19;
				aoa_ColumnSpliceBySignature(aoa, ENUM_ROW, '2', 0, 1); // добавить пустую колонку перед пунктом "2"
				aoa_ColumnSpliceBySignature(aoa, ENUM_ROW, '5б', 1, 2); // удалить 5б и добавить 2 колонки на её месте
				aoa_ColumnSpliceBySignature(aoa, ENUM_ROW, '8б', 0, 4); // вставить 4 колонки перед пунктом 8б

				for (let i = 0; i < OUTPUT_COL_NUM; i++) {
					// добавить нумерацию столбцов
					aoa[ENUM_ROW][i] = i + 1;
				}
				aoa.splice(2, 0, ['']); // добавить пустую сторку перед третьей
				aoa.splice(8, 2); // удалить две строки после 8-й
			}
		},
		sell: {
			output: '0000090 5_10 Книга продаж.xls',
			process: (/** @type {Array<any>} - array-of-arrays extracted from the file  */ aoa) => {
				const DATA_IS_PRESENT_COL = 0;

				// меняем формат ИИН и КПП
				const IIN_COL = 6,
					KPP_COL = 7;
				const DATA_STARTS_ROW = 11;
				for (let r = DATA_STARTS_ROW; r < aoa.length; r++) {
					if (aoa[r][DATA_IS_PRESENT_COL] != '') {
						aoa[r][IIN_COL] = aoa[r][IIN_COL] + '/' + aoa[r][KPP_COL];
					}
				}
				const ENUM_ROW = 10;
				aoa_ColumnSpliceBySignature(aoa, ENUM_ROW, '3а', 1, 4); // удалить 3а и добавить 4 колонки на её месте
				aoa_ColumnSpliceBySignature(aoa, ENUM_ROW, '5а', 0, 7); // вставить 7 колонок перед пунктом 5а
				aoa_CopyColumnBySignature(aoa, ENUM_ROW, '8а', 'row-13', DATA_STARTS_ROW); // расставляем колонки по местам
				aoa_CopyColumnBySignature(aoa, ENUM_ROW, '5а', 'row-14', DATA_STARTS_ROW);
				aoa_CopyColumnBySignature(aoa, ENUM_ROW, '6а', 'row-15', DATA_STARTS_ROW);
				aoa_CopyColumnBySignature(aoa, ENUM_ROW, '7', 'row-16', DATA_STARTS_ROW);
				aoa_CopyColumnBySignature(aoa, ENUM_ROW, '8б', 'row-17', DATA_STARTS_ROW);
				aoa_CopyColumnBySignature(aoa, ENUM_ROW, '5б', 'row-18', DATA_STARTS_ROW);
				aoa_CopyColumnBySignature(aoa, ENUM_ROW, '6б', 'row-19', DATA_STARTS_ROW);
				aoa_ColumnSpliceBySignature(aoa, ENUM_ROW, '5а', 0, 5); // добавить ещё 5 колонок
				aoa_ColumnSpliceBySignature(aoa, ENUM_ROW, '1', 0, 2, 'Продажа', 'Руководитель'); // сдвинуть всё между "Продажа" и "Руководитель" на две колонки
				aoa_EnumerateDataBySignature(aoa, ENUM_ROW, 'row-0', '1', DATA_STARTS_ROW); // добавить нумерацию строк
				const OUTPUT_COL_NUM = 27;
				for (let i = 0; i < OUTPUT_COL_NUM; i++) {
					// добавить нумерацию столбцов
					aoa[ENUM_ROW][i] = i + 1;
				}
				aoa.splice(2, 0, ['']); // добавить пустую сторку перед третьей
				aoa.splice(8, 2); // удалить две строки после 8-й
			}
		}
	}
];
