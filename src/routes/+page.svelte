<script lang="ts">
	import { read, utils, writeFile, type WorkBook } from 'xlsx';
	import {delete_rows, delete_cols}  from '$lib/helper.js';

	const schema = [
		{
			id: '5.09',
			buy: { output: '0000080 5_09 Книга покупок.xlsx', 'to-del': [11, 9, 'O'] },
			sell: { output: '0000090 5_09 Книга продаж.xlsx', 'to-del': [9, 7] }
		}
	];
	let selected = schema[0];
	let book: string = '';
	let processDisabled = true;

	let files: FileList;
	let excelData: WorkBook;
	let sheet;

	$: if (files && files[0]) {
		// Note that `files` is of type `FileList`, not an Array:
		// https://developer.mozilla.org/en-US/docs/Web/API/FileList
		let file = files[0];

		let reader = new FileReader();
		reader.onloadend = function (event) {
			let arrayBuffer = reader.result;
			excelData = read(arrayBuffer, { type: 'array' });
			//let sheet_id = excelData["Workbook"]["Sheets"][0]["name"];
			//let sheet = excelData["Sheets"][sheet_id];
		};

		reader.readAsArrayBuffer(file);
	}

	$: if (excelData) {
		let sheet_id = excelData['Workbook']['Sheets'][0]['name'];
		sheet = excelData['Sheets'][sheet_id];
		processDisabled = false;
		book = {'Книга покупок': 'buy', 'Книга продаж': 'sell'}[sheet['A2'].v] || ''

		console.log('DATA triggered', book);
	}

	function processFile() {
		function isNumber(value) {
			return typeof value == 'number';
		}


		let sheet_id = excelData['Workbook']['Sheets'][0]['name'];
	    let sheet = excelData['Sheets'][sheet_id];


		let sel_schema = selected[book];
		let to_del = sel_schema['to-del'];
		let output_name = sel_schema['output'];

		let rows_to_del = to_del.filter((x) => isNumber(x));
		let cols_to_del = to_del.filter((x) => !isNumber(x));
		console.log('COLS', cols_to_del, 'ROWS', rows_to_del);

		// return index of a column based on its name
		function alphaToNum(letters) {
			for (var p = 0, n = 0; p < letters.length; p++) {
				n = letters[p].charCodeAt() - 'A'.charCodeAt(0) + n * 26;
			}
			// console.log("AtoN", letters, n)
			return n;
		}
		
		for (let r in rows_to_del) {
			delete_rows(sheet, rows_to_del[r]);
		}

		for (let c in cols_to_del) {
			delete_cols(sheet, cols_to_del[c]);
		}

		writeFile(excelData, output_name);
	}
</script>

<section id="controls">
	<label for="selVer">Версия выходного документа</label>
	<select id="selVer" bind:value={selected}>
		{#each schema as version}
			<option value={version}>
				{version.id}
			</option>
		{/each}
	</select>

	<label for="file"
		>Выберите файл с данными выгрузки из 1С (Excel):
		<input id="file" type="file" accept=".xls,.xlsx" bind:files /></label>

		<label for="selBook">Формат книги:</label>
		<select id="selBook" bind:value={book} required>
			<option value="buy">Книга покупок</option>
			<option value="sell">Книга продаж</option>
		</select>

	<button on:click={() => processFile()} disabled={!book}>Обработка</button>
</section>
