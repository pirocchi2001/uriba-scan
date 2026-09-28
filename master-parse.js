/**
 * 転送.xlsx の全シートから D列=JAN, E列=品番名, F列=在売価（税込）を読み取る。
 * 見出し行は無視し、重複JANは先勝ち。戻り値: {JAN: [商品名, 在売価]}
 * ブラウザ（app.js）と暗号化ツール（tools/encrypt-master.js）の両方で使う。
 */
(function (root) {
  function parseMaster(XLSX, data) {
    const wb = XLSX.read(data, { type: 'array' });
    const items = {};
    for (const sheetName of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null });
      for (const r of rows) {
        if (!r || r[3] == null) continue;
        const jan = (typeof r[3] === 'number' ? String(Math.round(r[3])) : String(r[3])).trim();
        if (!/^\d{8,14}$/.test(jan)) continue; // 見出し行など
        if (items[jan]) continue;
        const name = r[4] == null ? '' : String(r[4]).trim();
        let price = null;
        if (typeof r[5] === 'number') price = Math.round(r[5]);
        else if (r[5] != null && /^\d+$/.test(String(r[5]).trim())) price = Number(String(r[5]).trim());
        items[jan] = [name, price];
      }
    }
    return items;
  }

  if (typeof module !== 'undefined' && module.exports) module.exports = parseMaster;
  else root.parseMaster = parseMaster;
})(this);
