<script lang="ts">
	import { read, utils, writeFile, type WorkBook } from 'xlsx';
	import { schema } from '../schema.js'

	let selected = schema[schema.length - 1];
	let book: string = '';
	$: book = book;

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
		selected.title_cells.forEach(function (title_cell) {
			console.log("CELL", title_cell)

			if (sheet[title_cell].v == "Книга покупок") {
				book = 'buy'
			}
			if (sheet[title_cell].v == "Книга продаж") {
				book = 'sell'
			}		
		})

		console.debug('DATA triggered', book);
	}

	function processFile() {
		let sheet_id = excelData['Workbook']['Sheets'][0]['name'];
	    let sheet = excelData['Sheets'][sheet_id];


		let sel_schema = selected[book];
		let output_name = sel_schema['output'];

		// return index of a column based on its name
		function alphaToNum(letters: String){
			var chrs = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ', mode = chrs.length - 1, number = 0;
			for(var p = 0; p < letters.length; p++){
				number = number * mode + chrs.indexOf(letters[p]);
			}
			return number-1;
		}

		const oneToZeroBased = (x: Number) => x-1;
		let rows_to_del = sel_schema['del-rows'].map(oneToZeroBased);
		let cols_to_del = sel_schema['del-cols'].map(alphaToNum);
		console.debug('COLS', cols_to_del, 'ROWS', rows_to_del);

		// deletes column from array-of-arrays
		function del_col_aoa(aoa, col_num: Number) {
			for (let i in aoa) {
				aoa[i].splice(col_num, 1);
			}
		}
		// deletes row from array-of-arrays
		function del_row_aoa(aoa, row_num: Number) {
			aoa.splice(row_num, 1);
		}

		let output = utils.book_new();
		let aoa = utils.sheet_to_json(sheet, { defval: "", header: 1 });

		for (let i in rows_to_del) {
			del_row_aoa(aoa, rows_to_del[i]);
		}

		for (let i in cols_to_del) {
			del_col_aoa(aoa, cols_to_del[i]);
		}

		let worksheet = utils.aoa_to_sheet(aoa);
		utils.book_append_sheet(output, worksheet, 'TDSheet');
		writeFile(output, output_name, {bookType: "biff8"});
	}
</script>

<section id="descr">
<p>Данное веб-приложение форматирует выгрузку книг продаж/покупок из 1С для последующей загрузки в СБИС. Обработка файлов происходит локально, никакие данные в интернет не загружаются.</p>
</section>

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

	<button on:click={() => processFile()}>Обработка</button>
</section>
<center><a href="https://github.com/The-Keeper/1c-to-sbis">Исходный код на GitHub</a></center>