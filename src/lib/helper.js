//@ts-nocheck
// === SOURCE: https://github.com/SheetJS/sheetjs/issues/413 ===

import { read, utils } from 'xlsx';

function clamp_range(range) {
	if(range.e.r >= (1<<20)) range.e.r = (1<<20)-1;
	if(range.e.c >= (1<<14)) range.e.c = (1<<14)-1;
	return range;
}

var crefregex = /(^|[^._A-Z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)([1-9]\d{0,5}|10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6])(?![_.\(A-Za-z0-9])/g;

/*
	inserts `nrows` rows BEFORE specified `start_row`
	- ws         = worksheet object
	- start_row  = starting row (0-indexed) | default 0
	- nrows      = number of rows to add    | default 1
	- opts       = options:
	  + fill     = set to true to "fill" cell styles and formulae
*/

function insert_rows(ws, start_row, nrows, opts) {
	if(!ws) throw new Error("operation expects a worksheet");
	var dense = Array.isArray(ws);
	if(!nrows) nrows = 1;
	if(!start_row) start_row = 0;
	if(!opts) opts = {};

	/* extract original range */
	var range = utils.decode_range(ws["!ref"]);
	var R = 0, C = 0;

	var formula_cb = function($0, $1, $2, $3, $4, $5) {
		var _R = utils.decode_row($5), _C = utils.decode_col($3);
		if(!opts.fill ? (_R >= start_row) : (R >= start_row)) _R += nrows;
		return $1+($2=="$" ? $2+$3 : utils.encode_col(_C))+($4=="$" ? $4+$5 : utils.encode_row(_R));
	};

	var addr, naddr, newcell;
	/* move cells and update formulae */
	if(dense) {
		/* cells after the insert */
		for(R = range.e.r; R >= start_row; --R) {
			if(ws[R]) ws[R].forEach(function(cell) { if(cell.f) cell.f = cell.f.replace(crefregex, formula_cb); });
			ws[R+nrows] = ws[R];
		}

		/* TODO: dense mode; newly created space */
		for(R = start_row; R < start_row + nrows; ++R) ws[R] = [];

		/* cells before insert */
		for(R = 0; R < start_row; ++R) {
			if(ws[R]) ws[R].forEach(function(cell) { if(cell.f) cell.f = cell.f.replace(crefregex, formula_cb); });
		}
		range.e.r += nrows;
	} else {
		/* cells after the insert */
		for(R = range.e.r; R >= start_row; --R) {
			for(C = range.s.c; C <= range.e.c; ++C) {
				addr = utils.encode_cell({r:R, c:C});
				naddr = utils.encode_cell({r:R+nrows, c:C});
				if(!ws[addr]) { delete ws[naddr]; continue; }
				if(opts.fill && (ws[addr].s || ws[addr].f)) {
					newcell = {};
					if(ws[addr].f) { newcell.f = ws[addr].f; newcell.t = ws[addr].t; }
					else { newcell.t = "z"; }
				}
				if(ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
				ws[naddr] = ws[addr];
				if(opts.fill) { ws[addr] = newcell; console.log(ws[addr], newcell); }
				if(range.e.r < R + nrows) range.e.r = R + nrows;
			}
		}

		/* newly created space */
		if(!opts.fill) for(R = start_row; R < start_row + nrows; ++R) {
			for(C = range.s.c; C <= range.e.c; ++C) {
				addr = utils.encode_cell({r:R, c:C});
				delete ws[addr];
			}
		}

		/* cells before insert */
		for(R = 0; R < start_row; ++R) {
			for(C = range.s.c; C <= range.e.c; ++C) {
				addr = utils.encode_cell({r:R, c:C});
				if(ws[addr] && ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
			}
		}
	}

	/* write new range */
	ws["!ref"] = utils.encode_range(clamp_range(range));

	/* merge cells */
	if(ws["!merges"]) ws["!merges"].forEach(function(merge, idx) {
		var mergerange;
		switch(typeof merge) {
			case 'string': mergerange = utils.decode_range(merge); break;
			case 'object': mergerange = merge; break;
			default: throw new Error("Unexpected merge ref " + merge);
		}
		if(mergerange.s.r >= start_row) mergerange.s.r += nrows;
		if(mergerange.e.r >= start_row) mergerange.e.r += nrows;
		clamp_range(mergerange);
		ws["!merges"][idx] = mergerange;
	});

	/* rows */
	var rowload = [start_row, 0];
	for(R = 0; R < nrows; ++R) rowload.push(void 0);
	if(ws["!rows"]) ws["!rows"].splice.apply(ws["!rows"], rowload);
}

/*
	deletes `nrows` rows STARTING WITH `start_row`
	- ws         = worksheet object
	- start_row  = starting row (0-indexed) | default 0
	- nrows      = number of rows to delete | default 1
*/

function delete_rows(ws, start_row, nrows) {
	if(!ws) throw new Error("operation expects a worksheet");
	var dense = Array.isArray(ws);
	if(!nrows) nrows = 1;
	if(!start_row) start_row = 0;

	/* extract original range */
	var range = utils.decode_range(ws["!ref"]);
	var R = 0, C = 0;

	var formula_cb = function($0, $1, $2, $3, $4, $5) {
		var _R = utils.decode_row($5), _C = utils.decode_col($3);
		if(_R >= start_row) {
			_R -= nrows;
			if(_R < start_row) return "#REF!";
		}
		return $1+($2=="$" ? $2+$3 : utils.encode_col(_C))+($4=="$" ? $4+$5 : utils.encode_row(_R));
	};

	var addr, naddr;
	/* move cells and update formulae */
	if(dense) {
		for(R = start_row + nrows; R <= range.e.r; ++R) {
			if(ws[R]) ws[R].forEach(function(cell) { cell.f = cell.f.replace(crefregex, formula_cb); });
			ws[R-nrows] = ws[R];
		}
		ws.length -= nrows;
		for(R = 0; R < start_row; ++R) {
			if(ws[R]) ws[R].forEach(function(cell) { cell.f = cell.f.replace(crefregex, formula_cb); });
		}
	} else {
		for(R = start_row + nrows; R <= range.e.r; ++R) {
			for(C = range.s.c; C <= range.e.c; ++C) {
				addr = utils.encode_cell({r:R, c:C});
				naddr = utils.encode_cell({r:R-nrows, c:C});
				if(!ws[addr]) { delete ws[naddr]; continue; }
				if(ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
				ws[naddr] = ws[addr];
			}
		}
		for(R = range.e.r; R > range.e.r - nrows; --R) {
			for(C = range.s.c; C <= range.e.c; ++C) {
				addr = utils.encode_cell({r:R, c:C});
				delete ws[addr];
			}
		}
		for(R = 0; R < start_row; ++R) {
			for(C = range.s.c; C <= range.e.c; ++C) {
				addr = utils.encode_cell({r:R, c:C});
				if(ws[addr] && ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
			}
		}
	}

	/* write new range */
	range.e.r -= nrows;
	if(range.e.r < range.s.r) range.e.r = range.s.r;
	ws["!ref"] = utils.encode_range(clamp_range(range));

	/* merge cells */
	if(ws["!merges"]) ws["!merges"].forEach(function(merge, idx) {
		var mergerange;
		switch(typeof merge) {
			case 'string': mergerange = utils.decode_range(merge); break;
			case 'object': mergerange = merge; break;
			default: throw new Error("Unexpected merge ref " + merge);
		}
		if(mergerange.s.r >= start_row) {
			mergerange.s.r = Math.max(mergerange.s.r - nrows, start_row);
			if(mergerange.e.r < start_row + nrows) { delete ws["!merges"][idx]; return; }
		} else if(mergerange.e.r >= start_row) mergerange.e.r = Math.max(mergerange.e.r - nrows, start_row);
		clamp_range(mergerange);
		ws["!merges"][idx] = mergerange;
	});
	if(ws["!merges"]) ws["!merges"] = ws["!merges"].filter(function(x) { return !!x; });

	/* rows */
	if(ws["!rows"]) ws["!rows"].splice(start_row, nrows);
}

/*
	inserts `ncols` cols BEFORE specified `start_col`
	- ws         = worksheet object
	- start_col  = starting col (0-indexed) | default 0
	- ncols      = number of cols to add    | default 1
	- opts       = options:
	  + fill     = set to true to "fill" cell styles and formulae
*/

function insert_cols(ws, start_col, ncols, opts) {
	if(!ws) throw new Error("operation expects a worksheet");
	var dense = Array.isArray(ws);
	if(!ncols) ncols = 1;
	if(!start_col) start_col = 0;
	if(!opts) opts = {};

	/* extract original range */
	var range = utils.decode_range(ws["!ref"]);
	var R = 0, C = 0;

	var formula_cb = function($0, $1, $2, $3, $4, $5) {
		var _R = utils.decode_row($5), _C = utils.decode_col($3);
		if(!opts.fill ? (_C >= start_col) : (C >= start_col)) _C += ncols;
		return $1+($2=="$" ? $2+$3 : utils.encode_col(_C))+($4=="$" ? $4+$5 : utils.encode_row(_R));
	};

	var addr, naddr, newcell;
	/* move cells and update formulae */
	if(dense) {
		for(R = range.s.r; R <= range.e.r; ++R) {
			if(!ws[R]) continue;
			/* cells before insert insert */
			for(C = 0; C < start_col; ++C) {
				if(ws[R][C] && ws[R][C].f) ws[R][C].f = ws[R][C].f.replace(crefregex, formula_cb);
			}
			/* cells after insert */
			for(C = range.e.c; C >= start_col; --C) {
				if(!ws[R][C]) { delete ws[R][C + ncols]; continue; }
				if(ws[R][C] && ws[R][C].f) ws[R][C].f = ws[R][C].f.replace(crefregex, formula_cb);
				ws[R][C + ncols] = ws[R][C];
				delete ws[R][C];
			}
		}
		range.e.c += ncols;
	} else {
		/* cells after the insert */
		for(C = range.e.c; C >= start_col; --C) {
			for(R = range.s.r; R <= range.e.r; ++R) {
				addr = utils.encode_cell({r:R, c:C});
				naddr = utils.encode_cell({r:R, c:C + ncols});
				if(!ws[addr]) { delete ws[naddr]; continue; }
				if(opts.fill && (ws[addr].s || ws[addr].f)) {
					newcell = {};
					if(ws[addr].f) { newcell.f = ws[addr].f; newcell.t = ws[addr].t; }
					else { newcell.t = "z"; }
				}
				if(ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
				ws[naddr] = ws[addr];
				if(opts.fill) { ws[addr] = newcell; console.log(ws[addr], newcell); }
				if(range.e.c < C + ncols) range.e.c = C + ncols;
			}
		}

		/* newly created space */
		if(!opts.fill) for(C = start_col; C < start_col + ncols; ++C) {
			for(R = range.s.r; R <= range.e.r; ++R) {
				addr = utils.encode_cell({r:R, c:C});
				delete ws[addr];
			}
		}

		/* cells before insert */
		for(C = 0; C < start_col; ++C) {
			for(R = range.s.r; R <= range.e.r; ++R) {
				addr = utils.encode_cell({r:R, c:C});
				if(ws[addr] && ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
			}
		}
	}

	/* write new range */
	ws["!ref"] = utils.encode_range(clamp_range(range));

	/* merge cells */
	if(ws["!merges"]) ws["!merges"].forEach(function(merge, idx) {
		var mergerange;
		switch(typeof merge) {
			case 'string': mergerange = utils.decode_range(merge); break;
			case 'object': mergerange = merge; break;
			default: throw new Error("Unexpected merge ref " + merge);
		}
		if(mergerange.s.c >= start_col) mergerange.s.c += ncols;
		if(mergerange.e.c >= start_col) mergerange.e.c += ncols;
		clamp_range(mergerange);
		ws["!merges"][idx] = mergerange;
	});

	/* cols */
	var colload = [start_col, 0];
	for(C = 0; C < ncols; ++C) colload.push(void 0);
	if(ws["!cols"]) ws["!cols"].splice.apply(ws["!cols"], colload);
}

/*
	deletes `ncols` cols STARTING WITH `start_col`
	- ws         = worksheet object
	- start_col  = starting col (0-indexed) | default 0
	- ncols      = number of cols to delete | default 1
*/

function delete_cols(ws, start_col, ncols) {
	if(!ws) throw new Error("operation expects a worksheet");
	var dense = Array.isArray(ws);
	if(!ncols) ncols = 1;
	if(!start_col) start_col = 0;

	/* extract original range */
	var range = utils.decode_range(ws["!ref"]);
	var R = 0, C = 0;

	var formula_cb = function($0, $1, $2, $3, $4, $5) {
		var _R = utils.decode_row($5), _C = utils.decode_col($3);
		if(_C >= start_col) {
			_C -= ncols;
			if(_C < start_col) return "#REF!";
		}
		return $1+($2=="$" ? $2+$3 : utils.encode_col(_C))+($4=="$" ? $4+$5 : utils.encode_row(_R));
	};

	var addr, naddr;
	/* move cells and update formulae */
	if(dense) {
	} else {
		for(C = start_col + ncols; C <= range.e.c; ++C) {
			for(R = range.s.r; R <= range.e.r; ++R) {
				addr = utils.encode_cell({r:R, c:C});
				naddr = utils.encode_cell({r:R, c:C - ncols});
				if(!ws[addr]) { delete ws[naddr]; continue; }
				if(ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
				ws[naddr] = ws[addr];
			}
		}
		for(C = range.e.c; C > range.e.c - ncols; --C) {
			for(R = range.s.r; R <= range.e.r; ++R) {
				addr = utils.encode_cell({r:R, c:C});
				delete ws[addr];
			}
		}
		for(C = 0; C < start_col; ++C) {
			for(R = range.s.r; R <= range.e.r; ++R) {
				addr = utils.encode_cell({r:R, c:C});
				if(ws[addr] && ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
			}
		}
	}

	/* write new range */
	range.e.c -= ncols;
	if(range.e.c < range.s.c) range.e.c = range.s.c;
	ws["!ref"] = utils.encode_range(clamp_range(range));

	/* merge cells */
	if(ws["!merges"]) ws["!merges"].forEach(function(merge, idx) {
		var mergerange;
		switch(typeof merge) {
			case 'string': mergerange = utils.decode_range(merge); break;
			case 'object': mergerange = merge; break;
			default: throw new Error("Unexpected merge ref " + merge);
		}
		if(mergerange.s.c >= start_col) {
			mergerange.s.c = Math.max(mergerange.s.c - ncols, start_col);
			if(mergerange.e.c < start_col + ncols) { delete ws["!merges"][idx]; return; }
			mergerange.e.c -= ncols;
			if(mergerange.e.c < mergerange.s.c) { delete ws["!merges"][idx]; return; }
		} else if(mergerange.e.c >= start_col) mergerange.e.c = Math.max(mergerange.e.c - ncols, start_col);
		clamp_range(mergerange);
		ws["!merges"][idx] = mergerange;
	});
	if(ws["!merges"]) ws["!merges"] = ws["!merges"].filter(function(x) { return !!x; });

	/* cols */
	if(ws["!cols"]) ws["!cols"].splice(start_col, ncols);
}

export {delete_cols, delete_rows}
