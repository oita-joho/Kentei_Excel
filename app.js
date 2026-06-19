let rules = [];
let results = [];

document.getElementById("gradeBtn").addEventListener("click", gradeAll);
document.getElementById("csvBtn").addEventListener("click", downloadCSV);

function setSummary(msg) {
  document.getElementById("summary").textContent = msg;
}

async function gradeAll() {
  try {
    const ruleFile = document.getElementById("ruleFile").files[0];
    const excelFiles = document.getElementById("excelFiles").files;

    if (!ruleFile) {
      alert("採点基準CSVを選択してください。");
      return;
    }

    if (excelFiles.length === 0) {
      alert("生徒のExcelファイルを選択してください。");
      return;
    }

    setSummary("採点基準CSVを読み込み中...");

    rules = await readRuleCSV(ruleFile);

    if (rules.length === 0) {
      alert("採点基準CSVが読み込めませんでした。");
      return;
    }

    setSummary(`採点基準 ${rules.length} 件を読み込みました。Excelを採点中...`);

    results = [];

    for (const file of excelFiles) {
      const result = await gradeExcel(file);
      results.push(result);
    }

    showResults();

  } catch (err) {
    console.error(err);
    setSummary("エラー：" + err.message);
    alert("エラーが出ました。画面下の採点結果欄を確認してください。");
  }
}

function readRuleCSV(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      let text = e.target.result.replace(/^\uFEFF/, "");
      const lines = text.trim().split(/\r?\n/);
      const data = [];

      for (let i = 1; i < lines.length; i++) {
        const cols = lines[i].split(",");

        if (cols.length < 6) continue;

        data.push({
          name: cols[0].trim(),
          sheet: cols[1].trim(),
          cell: cols[2].trim(),
          type: cols[3].trim(),
          answer: cols[4].trim(),
          point: Number(cols[5])
        });
      }

      resolve(data);
    };

    reader.readAsText(file, "UTF-8");
  });
}

function gradeExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array", cellFormula: true });

        let score = 0;
        let max = 0;
        const details = [];

        for (const rule of rules) {
          max += rule.point;

          const sheet = workbook.Sheets[rule.sheet];

          if (!sheet) {
            details.push(
              `${rule.name}: × シート「${rule.sheet}」なし（実際のシート：${workbook.SheetNames.join("、")}）`
            );
            continue;
          }

          const cell = sheet[rule.cell];

          if (!cell) {
            details.push(`${rule.name}: × セル「${rule.cell}」なし`);
            continue;
          }

          let ok = false;

          if (rule.type === "value") {
            ok = String(cell.v ?? "").trim() === String(rule.answer).trim();
          }

          if (rule.type === "formula") {
            const formula = String(cell.f ?? "").toUpperCase();
            ok = formula.includes(String(rule.answer).toUpperCase());
          }

          if (ok) {
            score += rule.point;
            details.push(`${rule.name}: ○ +${rule.point}`);
          } else {
            details.push(`${rule.name}: × 実際の値=${cell.v ?? ""} 数式=${cell.f ?? ""}`);
          }
        }

        resolve({
          fileName: file.name,
          score,
          max,
          details: details.join(" / ")
        });

      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = () => reject(new Error("Excelファイルを読み込めませんでした。"));
    reader.readAsArrayBuffer(file);
  });
}

function showResults() {
  const tbody = document.getElementById("resultBody");
  const summary = document.getElementById("summary");

  tbody.innerHTML = "";

  for (const r of results) {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${escapeHtml(r.fileName)}</td>
      <td>${r.score}</td>
      <td>${r.max}</td>
      <td>${escapeHtml(r.details)}</td>
    `;

    tbody.appendChild(tr);
  }

  summary.textContent = `${results.length}件のファイルを採点しました。`;
  document.getElementById("csvBtn").disabled = false;
}

function downloadCSV() {
  const rows = [["ファイル名", "得点", "満点", "詳細"]];

  for (const r of results) {
    rows.push([r.fileName, r.score, r.max, r.details]);
  }

  const csv = rows.map(row =>
    row.map(v => `"${String(v).replaceAll('"', '""')}"`).join(",")
  ).join("\n");

  const blob = new Blob(["\uFEFF" + csv], {
    type: "text/csv;charset=utf-8;"
  });

  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "採点結果.csv";
  a.click();
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}
