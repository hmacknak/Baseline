// Office.js stand-in for the demo video. The add-in runs unmodified; Excel calls are answered
// from the simulated workbook in the parent page (studio.html).
(function () {
  const P = window.parent;
  const LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  function extents(sheet) {
    let rows = 0, cols = 0;
    Object.keys(sheet.cells).forEach(a => {
      const m = /^([A-Z]+)(\d+)$/.exec(a);
      rows = Math.max(rows, +m[2]);
      cols = Math.max(cols, LETTERS.indexOf(m[1]) + 1);
    });
    return { rows, cols };
  }

  function typeOf(v) {
    if (v === null || v === undefined || v === "") return "Empty";
    if (typeof v === "number") return "Double";
    if (typeof v === "boolean") return "Boolean";
    if (typeof v === "string" && v[0] === "#") return "Error";
    return "String";
  }

  function wsFor(name) {
    const sheet = () => P.__wb.sheets[name];
    return {
      name,
      load() {},
      getUsedRangeOrNullObject() {
        const e = extents(sheet());
        return { load() {}, isNullObject: e.rows === 0, rowIndex: 0, columnIndex: 0, rowCount: e.rows, columnCount: e.cols,
          get values() { return grid(name, e.rows, e.cols).values; } };
      },
      getRangeByIndexes(r0, c0, rows, cols) {
        const g = grid(name, rows, cols);
        return {
          load() {},
          values: g.values,
          formulas: g.formulas,
          valueTypes: g.types,
          numberFormat: g.formats,
          getCellProperties() {
            return { value: g.fills.map(row => row.map(f => ({ format: { fill: { color: f || "#FFFFFF" } } }))) };
          },
        };
      },
      visibility: "Visible",
      activate() { P.__activate(name); },
      getRange(a) { return { select() { P.__select(a); } }; },
    };
  }

  function grid(name, rows, cols) {
    const s = P.__wb.sheets[name];
    const values = [], formulas = [], types = [], formats = [], fills = [];
    for (let r = 0; r < rows; r++) {
      values.push([]); formulas.push([]); types.push([]); formats.push([]); fills.push([]);
      for (let c = 0; c < cols; c++) {
        const a = LETTERS[c] + (r + 1);
        const cell = s.cells[a];
        const raw = cell ? cell.v : "";
        const val = cell ? P.__valueOf(name, a) : "";
        values[r].push(val === null ? "" : val);
        formulas[r].push(raw);
        types[r].push(typeOf(val));
        formats[r].push(cell && cell.fmt ? cell.fmt : "General");
        fills[r].push(cell && cell.fill ? cell.fill : null);
      }
    }
    return { values, formulas, types, formats, fills };
  }

  const worksheets = {
    getActiveWorksheet() { return wsFor(P.__wb.active); },
    getItem(name) { return wsFor(name); },
    onChanged: { add(h) { P.__changeHandlers.push(h); } },
    onActivated: { add(h) { P.__activateHandlers.push(h); } },
    load() {},
    get items() { return P.__wb.order.map(wsFor); },
  };

  const settings = {};
  window.Excel = {
    SheetVisibility: { visible: "Visible" },
    run: async cb => cb({ workbook: { worksheets }, sync: async () => {} }),
  };
  window.Office = {
    AsyncResultStatus: { Succeeded: "succeeded" },
    onReady(cb) { const go = () => (P.__wb ? cb() : setTimeout(go, 50)); setTimeout(go, 0); return Promise.resolve(); },
    context: { document: { settings: {
      get: k => settings[k],
      set: (k, v) => { settings[k] = JSON.parse(JSON.stringify(v)); },
      saveAsync: cb => cb({ status: "succeeded" }),
    } } },
  };
})();
