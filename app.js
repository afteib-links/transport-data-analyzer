const runButton = document.getElementById("runButton");
const generateMasterButton = document.getElementById("generateMasterButton");

const excelInput = document.getElementById("excelFiles");
const companyMasterInput = document.getElementById("companyMaster");
const projectMasterInput = document.getElementById("projectMaster");
const companySourceMasterInput = document.getElementById("companySourceMaster");
const projectSourceMasterInput = document.getElementById("projectSourceMaster");

const logEl = document.getElementById("log");
const downloadsEl = document.getElementById("downloads");
const masterDownloadsEl = document.getElementById("masterDownloads");
const masterStatusEl = document.getElementById("masterStatus");

const SHEET_KEYWORDS = {
  db: ["DB"],
  invoice: ["請求"],
  payout: ["支払"],
};

const DEDUCTION_LABELS = ["手数料", "会費", "車両代", "調整"];

const COMPANY_MASTER_HEADERS = [
  "担当",
  "形態",
  "検索用",
  "企業No",
  "事業所No",
  "企業名",
  "カナ",
  "郵便番号",
  "請求書送付先住所",
  "電話",
  "FAX",
  "車両番号",
  "車検証有効期限",
  "任意保険有効期限",
  "締日",
  "支払日",
  "基本契約日",
  "業務内容及び付帯業務",
  "銀行名",
  "支店名",
  "口座番号",
  "預金名",
  "口座名義",
  "契約担当者",
  "現場担当者",
];

const PROJECT_MASTER_HEADERS = [
  "担当",
  "締日",
  "形態",
  "検索用",
  "氏名",
  "稼働企業",
  "企業No",
  "事業所No",
  "支払区分",
  "分割単価",
  "郵便番号",
  "住所",
  "電話",
  "生年月日",
  "年齢",
  "血液",
  "振込口座",
  "振込支店",
  "口座種類",
  "口座番号",
  "口座名義",
  "支払日",
  "委託単価",
  "契約単価",
  "インボイス番号",
  "事務手数料",
  "安全協力会費",
  "傷害",
  "請負損害",
  "貨物",
  "Ｇ会",
  "確定申告",
  "過去安全大会",
  "稼働開始日",
  "基本契約日",
  "稼働履歴",
  "継続年数",
  "免許有効期限",
  "車両番号",
  "車検有効期限",
  "任意保険期限",
  "ループ",
  "フリガナ（氏）",
  "パートナーフリ",
];

const generatedMasterState = {
  companyRows: [],
  projectRows: [],
};

generateMasterButton.addEventListener("click", async () => {
  try {
    const companyFile = companySourceMasterInput.files?.[0];
    const projectFile = projectSourceMasterInput.files?.[0];

    if (!companyFile && !projectFile) {
      setMasterStatus("企業または案件の元データを選択してください。", true);
      return;
    }

    const loadedCompanyRows = companyFile ? await readDelimitedFile(companyFile) : [];
    const loadedProjectRows = projectFile ? await readDelimitedFile(projectFile) : [];

    let companyRows = normalizeRowsToHeaders(loadedCompanyRows, COMPANY_MASTER_HEADERS);
    let projectRows = normalizeRowsToHeaders(loadedProjectRows, PROJECT_MASTER_HEADERS);

    const result = assignNumbersAndLink(companyRows, projectRows);
    companyRows = result.companyRows;
    projectRows = result.projectRows;

    generatedMasterState.companyRows = companyRows;
    generatedMasterState.projectRows = projectRows;

    masterDownloadsEl.innerHTML = "";
    renderDownload("company_master_generated.csv", toCsv(companyRows, COMPANY_MASTER_HEADERS), masterDownloadsEl);
    renderDownload("project_master_generated.csv", toCsv(projectRows, PROJECT_MASTER_HEADERS), masterDownloadsEl);

    setMasterStatus(
      `生成完了: 企業 ${companyRows.length} 件 / 案件 ${projectRows.length} 件（新規企業No ${result.newCompanyNoCount} 件, 新規事業所No ${result.newOfficeNoCount} 件）`,
      false
    );
    appendLog("マスター生成完了: 生成データを解析時のマスターとして自動利用します。");
  } catch (error) {
    setMasterStatus(`エラー: ${error.message}`, true);
  }
});

runButton.addEventListener("click", async () => {
  if (!window.XLSX) {
    appendLog("エラー: xlsx.full.min.js が見つかりません。");
    return;
  }

  const excelFiles = Array.from(excelInput.files || []).filter((file) =>
    file.name.toLowerCase().endsWith(".xlsx")
  );

  if (excelFiles.length === 0) {
    appendLog("エラー: .xlsx ファイルが選択されていません。");
    return;
  }

  setRunning(true);
  clearOutput();

  try {
    const uploadedCompanyRows = await readOptionalMasterRows(companyMasterInput.files?.[0], COMPANY_MASTER_HEADERS);
    const uploadedProjectRows = await readOptionalMasterRows(projectMasterInput.files?.[0], PROJECT_MASTER_HEADERS);

    const activeCompanyRows = generatedMasterState.companyRows.length > 0
      ? generatedMasterState.companyRows
      : uploadedCompanyRows;
    const activeProjectRows = generatedMasterState.projectRows.length > 0
      ? generatedMasterState.projectRows
      : uploadedProjectRows;

    appendLog(`開始: ${excelFiles.length} ファイルを解析します。`);

    const records = {
      daily: [],
      pl: [],
      invoice: [],
      payout: [],
    };

    for (const file of excelFiles) {
      const meta = extractPathMeta(file.webkitRelativePath || file.name);
      const workbook = await readWorkbook(file);

      const dbSheets = workbook.SheetNames.filter((name) =>
        includesAnyKeyword(name, SHEET_KEYWORDS.db)
      );
      const invoiceSheets = workbook.SheetNames.filter((name) =>
        includesAnyKeyword(name, SHEET_KEYWORDS.invoice)
      );
      const payoutSheets = workbook.SheetNames.filter((name) =>
        includesAnyKeyword(name, SHEET_KEYWORDS.payout)
      );

      for (const sheetName of dbSheets) {
        const sheet = workbook.Sheets[sheetName];
        const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
        const dbData = extractDbSheetData(matrix);

        const companyMasterRow = findBestCompanyRow(activeCompanyRows, meta.companyName) || {};
        const projectMasterRow = findBestProjectRow(activeProjectRows, dbData.partnerName) || {};

        records.daily.push(
          ...dbData.dailyRows.map((row) => ({
            accountingPeriod: meta.accountingPeriod,
            groupName: meta.groupName,
            companyName: meta.companyName,
            sourceFile: meta.fileName,
            sourceSheet: sheetName,
            partnerName: dbData.partnerName,
            ...row,
            companyNo: companyMasterRow["企業No"] || projectMasterRow["企業No"] || "",
            officeNo: companyMasterRow["事業所No"] || projectMasterRow["事業所No"] || "",
            closingDay: companyMasterRow["締日"] || projectMasterRow["締日"] || "",
            paymentDay: companyMasterRow["支払日"] || projectMasterRow["支払日"] || "",
            invoiceNumber: projectMasterRow["インボイス番号"] || "",
            contractUnitPrice: projectMasterRow["契約単価"] || "",
            officeFeeMaster: projectMasterRow["事務手数料"] || "",
          }))
        );

        records.pl.push({
          accountingPeriod: meta.accountingPeriod,
          groupName: meta.groupName,
          companyName: meta.companyName,
          sourceFile: meta.fileName,
          sourceSheet: sheetName,
          partnerName: dbData.partnerName,
          month: detectMonth(sheetName),
          totalSales: dbData.totalSales,
          totalPayment: dbData.totalPayment,
          actualPayment: dbData.actualPayment,
          grossProfit: dbData.grossProfit,
          fee: dbData.deductions["手数料"] || "",
          membershipFee: dbData.deductions["会費"] || "",
          vehicleCost: dbData.deductions["車両代"] || "",
          adjustment: dbData.deductions["調整"] || "",
          hasError: "",
          errorMessage: "",
        });
      }

      for (const sheetName of invoiceSheets) {
        const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
          header: 1,
          raw: false,
          defval: "",
        });
        const invoiceData = extractInvoiceSheetData(matrix);
        records.invoice.push({
          accountingPeriod: meta.accountingPeriod,
          groupName: meta.groupName,
          companyName: meta.companyName,
          sourceFile: meta.fileName,
          sourceSheet: sheetName,
          month: detectMonth(sheetName),
          partnerName: invoiceData.partnerName,
          destination: invoiceData.destination,
          totalAmount: invoiceData.totalAmount,
        });
      }

      for (const sheetName of payoutSheets) {
        const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
          header: 1,
          raw: false,
          defval: "",
        });
        const payoutData = extractPayoutSheetData(matrix);
        records.payout.push({
          accountingPeriod: meta.accountingPeriod,
          groupName: meta.groupName,
          companyName: meta.companyName,
          sourceFile: meta.fileName,
          sourceSheet: sheetName,
          month: detectMonth(sheetName),
          partnerName: payoutData.partnerName,
          scheduledTransferDate: payoutData.scheduledTransferDate,
          amount: payoutData.amount,
          invoiceNumber: payoutData.invoiceNumber,
        });
      }

      appendLog(`解析済み: ${meta.fileName}`);
    }

    validateConsistency(records);

    const dailyCsv = toCsv(records.daily, [
      "accountingPeriod",
      "groupName",
      "companyName",
      "partnerName",
      "date",
      "startTime",
      "endTime",
      "excessCount",
      "attendanceDays",
      "remarks",
      "companyNo",
      "officeNo",
      "closingDay",
      "paymentDay",
      "invoiceNumber",
      "contractUnitPrice",
      "officeFeeMaster",
      "sourceFile",
      "sourceSheet",
    ]);

    const plCsv = toCsv(records.pl, [
      "accountingPeriod",
      "groupName",
      "companyName",
      "partnerName",
      "month",
      "totalSales",
      "totalPayment",
      "actualPayment",
      "grossProfit",
      "fee",
      "membershipFee",
      "vehicleCost",
      "adjustment",
      "hasError",
      "errorMessage",
      "sourceFile",
      "sourceSheet",
    ]);

    const invoiceCsv = toCsv(records.invoice, [
      "accountingPeriod",
      "groupName",
      "companyName",
      "partnerName",
      "month",
      "destination",
      "totalAmount",
      "sourceFile",
      "sourceSheet",
    ]);

    const payoutCsv = toCsv(records.payout, [
      "accountingPeriod",
      "groupName",
      "companyName",
      "partnerName",
      "month",
      "scheduledTransferDate",
      "amount",
      "invoiceNumber",
      "sourceFile",
      "sourceSheet",
    ]);

    renderDownload("daily_data.csv", dailyCsv, downloadsEl);
    renderDownload("pl_data.csv", plCsv, downloadsEl);
    renderDownload("invoice_summary.csv", invoiceCsv, downloadsEl);
    renderDownload("payout_summary.csv", payoutCsv, downloadsEl);

    appendLog("完了: 4件のCSVを生成しました。");
  } catch (error) {
    appendLog(`エラー: ${error.message}`);
    console.error(error);
  } finally {
    setRunning(false);
  }
});

function setRunning(running) {
  runButton.disabled = running;
  runButton.textContent = running ? "解析中..." : "解析してCSVを生成";
}

function setMasterStatus(message, isError) {
  masterStatusEl.textContent = message;
  masterStatusEl.style.color = isError ? "#b91c1c" : "#047857";
}

function appendLog(message) {
  const timestamp = new Date().toLocaleTimeString("ja-JP");
  logEl.textContent += `[${timestamp}] ${message}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}

function clearOutput() {
  logEl.textContent = "";
  downloadsEl.innerHTML = "";
}

function renderDownload(fileName, csvText, container) {
  const bom = "\uFEFF";
  const blob = new Blob([bom + csvText], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  const li = document.createElement("li");
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.textContent = fileName;
  li.appendChild(a);
  container.appendChild(li);
}

function includesAnyKeyword(text, keywords) {
  return keywords.some((keyword) => String(text).includes(keyword));
}

function normalize(value) {
  return String(value || "").replace(/\s+/g, "").trim();
}

function normalizeForMatch(value) {
  return String(value || "")
    .replace(/株式会社|有限会社|合同会社|合資会社|合名会社|（株）|\(株\)|（有）|\(有\)|[\s　]|ー|-|－|[(（)）]/g, "")
    .trim();
}

function isEmpty(value) {
  return normalize(value) === "";
}

function findLabelCell(matrix, labels) {
  const normalizedLabels = labels.map((label) => normalize(label));
  for (let rowIndex = 0; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] || [];
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const cell = normalize(row[colIndex]);
      if (!cell) {
        continue;
      }
      if (normalizedLabels.some((label) => cell.includes(label))) {
        return { rowIndex, colIndex };
      }
    }
  }
  return null;
}

function findRightValue(matrix, rowIndex, colIndex) {
  const row = matrix[rowIndex] || [];
  for (let col = colIndex + 1; col < row.length; col += 1) {
    if (!isEmpty(row[col])) {
      return row[col];
    }
  }
  return "";
}

function parseNumber(value) {
  const normalized = String(value || "").replace(/,/g, "").trim();
  if (!normalized) {
    return "";
  }
  const num = Number(normalized);
  return Number.isFinite(num) ? num : "";
}

function extractAnchorValue(matrix, labels) {
  const cell = findLabelCell(matrix, labels);
  if (!cell) {
    return "";
  }
  return findRightValue(matrix, cell.rowIndex, cell.colIndex);
}

function detectMonth(sheetName) {
  const matched = String(sheetName).match(/(\d{1,2}\.\d{1,2}|\d{1,2}\/\d{1,2}|\d{2,4}-\d{1,2})/);
  return matched ? matched[1] : "";
}

function extractPathMeta(path) {
  const segments = String(path).split("/").filter(Boolean);
  const fileName = segments[segments.length - 1] || "";
  const [accountingPeriod = "", groupName = "", companyName = ""] = segments.slice(-4, -1);
  return {
    accountingPeriod,
    groupName,
    companyName,
    fileName,
  };
}

function findHeaderRow(matrix, labels) {
  for (let rowIndex = 0; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] || [];
    const cells = row.map((cell) => normalize(cell));
    if (labels.every((label) => cells.some((cell) => cell.includes(normalize(label))))) {
      return { rowIndex, row };
    }
  }
  return null;
}

function findColumnIndexByKeywords(row, labels) {
  for (let index = 0; index < row.length; index += 1) {
    const value = normalize(row[index]);
    if (!value) {
      continue;
    }
    if (labels.some((label) => value.includes(normalize(label)))) {
      return index;
    }
  }
  return -1;
}

function extractDbSheetData(matrix) {
  const partnerName = extractAnchorValue(matrix, ["氏名", "パートナー", "乗務員"]);
  const totalSales = parseNumber(extractAnchorValue(matrix, ["総売上額", "売上合計"]));
  const totalPayment = parseNumber(extractAnchorValue(matrix, ["総支払額", "支払合計"]));
  const actualPayment = parseNumber(extractAnchorValue(matrix, ["実際支払額", "実支払額"]));
  const grossProfit = parseNumber(extractAnchorValue(matrix, ["粗利"]));

  const deductions = {};
  for (const label of DEDUCTION_LABELS) {
    deductions[label] = parseNumber(extractAnchorValue(matrix, [label]));
  }

  const header = findHeaderRow(matrix, ["始業", "終業"]);
  const dailyRows = [];

  if (header) {
    const dateCol = findColumnIndexByKeywords(header.row, ["日付", "日", "稼働日"]);
    const startCol = findColumnIndexByKeywords(header.row, ["始業", "開始"]);
    const endCol = findColumnIndexByKeywords(header.row, ["終業", "終了"]);
    const excessCol = findColumnIndexByKeywords(header.row, ["超過"]);
    const attendanceCol = findColumnIndexByKeywords(header.row, ["出勤", "日数"]);
    const remarksCol = findColumnIndexByKeywords(header.row, ["備考", "メモ"]);

    for (let rowIndex = header.rowIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
      const row = matrix[rowIndex] || [];
      const rowText = normalize(row.join(""));

      if (!rowText) {
        continue;
      }
      if (["合計", "総計", "計"].some((keyword) => rowText.includes(keyword))) {
        break;
      }

      const date = dateCol >= 0 ? row[dateCol] : "";
      const startTime = startCol >= 0 ? row[startCol] : "";
      const endTime = endCol >= 0 ? row[endCol] : "";
      const excessCount = excessCol >= 0 ? row[excessCol] : "";
      const attendanceDays = attendanceCol >= 0 ? row[attendanceCol] : "";
      const remarks = remarksCol >= 0 ? row[remarksCol] : "";

      if ([date, startTime, endTime, excessCount, attendanceDays, remarks].every(isEmpty)) {
        continue;
      }

      dailyRows.push({
        date,
        startTime,
        endTime,
        excessCount,
        attendanceDays,
        remarks,
      });
    }
  }

  return {
    partnerName,
    totalSales,
    totalPayment,
    actualPayment,
    grossProfit,
    deductions,
    dailyRows,
  };
}

function extractInvoiceSheetData(matrix) {
  return {
    partnerName: extractAnchorValue(matrix, ["氏名", "パートナー", "乗務員"]),
    destination: extractAnchorValue(matrix, ["宛先", "請求先", "御中"]),
    totalAmount: parseNumber(extractAnchorValue(matrix, ["合計", "請求合計", "総請求額"])),
  };
}

function extractPayoutSheetData(matrix) {
  return {
    partnerName: extractAnchorValue(matrix, ["氏名", "パートナー", "乗務員"]),
    scheduledTransferDate: extractAnchorValue(matrix, ["振込予定日", "支払予定日"]),
    amount: parseNumber(extractAnchorValue(matrix, ["お支払金額", "支払金額", "合計"])),
    invoiceNumber: extractAnchorValue(matrix, ["インボイス番号", "登録番号"]),
  };
}

function validateConsistency(records) {
  const invoiceMap = new Map();
  for (const row of records.invoice) {
    const key = [row.companyName, row.partnerName, row.month].join("|");
    invoiceMap.set(key, row.totalAmount);
  }

  const payoutMap = new Map();
  for (const row of records.payout) {
    const key = [row.companyName, row.partnerName, row.month].join("|");
    payoutMap.set(key, row.amount);
  }

  for (const row of records.pl) {
    const key = [row.companyName, row.partnerName, row.month].join("|");
    const invoiceAmount = parseNumber(invoiceMap.get(key));
    const payoutAmount = parseNumber(payoutMap.get(key));

    const errors = [];
    if (invoiceAmount !== "" && row.totalSales !== "" && invoiceAmount !== row.totalSales) {
      errors.push("DB総売上と請求合計が不一致");
    }
    if (payoutAmount !== "" && row.actualPayment !== "" && payoutAmount !== row.actualPayment) {
      errors.push("DB実際支払額と支払金額が不一致");
    }

    row.hasError = errors.length > 0 ? "1" : "0";
    row.errorMessage = errors.join("; ");

    if (errors.length > 0) {
      appendLog(`警告: ${row.companyName}/${row.partnerName}/${row.month} -> ${row.errorMessage}`);
    }
  }
}

function findBestCompanyRow(companyRows, inputName) {
  return findBestByName(companyRows, inputName, ["企業名", "検索用"]);
}

function findBestProjectRow(projectRows, inputName) {
  return findBestByName(projectRows, inputName, ["氏名", "検索用", "パートナーフリ"]);
}

function findBestByName(rows, inputName, fields) {
  const normalizedInput = normalizeForMatch(inputName);
  if (!normalizedInput || rows.length === 0) {
    return null;
  }

  for (const row of rows) {
    for (const field of fields) {
      const candidate = normalizeForMatch(row[field]);
      if (candidate && candidate === normalizedInput) {
        return row;
      }
    }
  }

  let best = null;
  let bestScore = 0;
  for (const row of rows) {
    for (const field of fields) {
      const candidate = normalizeForMatch(row[field]);
      if (!candidate) {
        continue;
      }
      const score = calculateSimilarity(normalizedInput, candidate);
      if (score > bestScore) {
        best = row;
        bestScore = score;
      }
    }
  }

  return bestScore >= 0.5 ? best : null;
}

function calculateSimilarity(s1, s2) {
  const len1 = s1.length;
  const len2 = s2.length;

  if (len1 === 0 && len2 === 0) {
    return 1;
  }
  if (len1 === 0 || len2 === 0) {
    return 0;
  }

  const matrix = [];
  for (let i = 0; i <= len1; i += 1) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= len2; j += 1) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= len1; i += 1) {
    for (let j = 1; j <= len2; j += 1) {
      const cost = s1[i - 1] === s2[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost
      );
    }
  }

  const distance = matrix[len1][len2];
  return 1 - distance / Math.max(len1, len2);
}

function assignNumbersAndLink(companyRows, projectRows) {
  const mutableCompanies = companyRows.map((row) => ({ ...row }));
  const mutableProjects = projectRows.map((row) => ({ ...row }));

  let newCompanyNoCount = 0;
  let newOfficeNoCount = 0;

  let nextCompany = getNextCompanyNumber([...mutableCompanies, ...mutableProjects]);
  const officeCounter = buildOfficeCounter([...mutableCompanies, ...mutableProjects]);

  for (const companyRow of mutableCompanies) {
    if (isEmpty(companyRow["企業No"])) {
      companyRow["企業No"] = formatCompanyNo(nextCompany);
      nextCompany += 1;
      newCompanyNoCount += 1;
    }
    if (isEmpty(companyRow["事業所No"])) {
      const companyNo = companyRow["企業No"];
      companyRow["事業所No"] = formatOfficeNo(nextOfficeNumber(officeCounter, companyNo));
      newOfficeNoCount += 1;
    } else {
      registerExistingOffice(officeCounter, companyRow["企業No"], companyRow["事業所No"]);
    }
  }

  const companyMatcherRows = mutableCompanies.map((row) => ({
    companyNo: row["企業No"],
    officeNo: row["事業所No"],
    companyName: row["企業名"],
    searchText: row["検索用"] || row["企業名"],
  }));

  for (const projectRow of mutableProjects) {
    const workingCompany = projectRow["稼働企業"] || "";
    let companyNo = projectRow["企業No"] || "";
    let officeNo = projectRow["事業所No"] || "";

    if (!companyNo) {
      const matched = findBestByName(
        companyMatcherRows.map((row) => ({
          企業名: row.companyName,
          検索用: row.searchText,
          企業No: row.companyNo,
          事業所No: row.officeNo,
        })),
        workingCompany,
        ["企業名", "検索用"]
      );

      if (matched) {
        companyNo = matched["企業No"] || "";
        if (!officeNo) {
          officeNo = matched["事業所No"] || "";
        }
      }
    }

    if (!companyNo) {
      companyNo = formatCompanyNo(nextCompany);
      nextCompany += 1;
      newCompanyNoCount += 1;

      const newCompany = createEmptyRow(COMPANY_MASTER_HEADERS);
      newCompany["企業No"] = companyNo;
      newCompany["事業所No"] = formatOfficeNo(nextOfficeNumber(officeCounter, companyNo));
      newCompany["企業名"] = workingCompany || projectRow["氏名"] || "未設定企業";
      newCompany["検索用"] = normalizeForMatch(newCompany["企業名"]);

      mutableCompanies.push(newCompany);
      companyMatcherRows.push({
        companyNo: newCompany["企業No"],
        officeNo: newCompany["事業所No"],
        companyName: newCompany["企業名"],
        searchText: newCompany["検索用"],
      });
      newOfficeNoCount += 1;
    }

    if (!officeNo) {
      officeNo = formatOfficeNo(nextOfficeNumber(officeCounter, companyNo));
      newOfficeNoCount += 1;
    } else {
      registerExistingOffice(officeCounter, companyNo, officeNo);
    }

    projectRow["企業No"] = companyNo;
    projectRow["事業所No"] = officeNo;
  }

  return {
    companyRows: mutableCompanies,
    projectRows: mutableProjects,
    newCompanyNoCount,
    newOfficeNoCount,
  };
}

function getNextCompanyNumber(rows) {
  let max = 0;
  for (const row of rows) {
    const value = row["企業No"];
    const n = extractTrailingNumber(value);
    if (n > max) {
      max = n;
    }
  }
  return max + 1;
}

function buildOfficeCounter(rows) {
  const map = new Map();
  for (const row of rows) {
    registerExistingOffice(map, row["企業No"], row["事業所No"]);
  }
  return map;
}

function registerExistingOffice(counterMap, companyNo, officeNo) {
  const c = String(companyNo || "").trim();
  const o = String(officeNo || "").trim();
  if (!c || !o) {
    return;
  }
  const value = extractTrailingNumber(o);
  const current = counterMap.get(c) || 0;
  if (value > current) {
    counterMap.set(c, value);
  }
}

function nextOfficeNumber(counterMap, companyNo) {
  const key = String(companyNo || "").trim();
  const current = counterMap.get(key) || 0;
  const next = current + 1;
  counterMap.set(key, next);
  return next;
}

function extractTrailingNumber(value) {
  const matched = String(value || "").match(/(\d+)$/);
  return matched ? Number(matched[1]) : 0;
}

function formatCompanyNo(number) {
  return `C${String(number).padStart(4, "0")}`;
}

function formatOfficeNo(number) {
  return `Z${String(number).padStart(3, "0")}`;
}

function createEmptyRow(headers) {
  const row = {};
  for (const header of headers) {
    row[header] = "";
  }
  return row;
}

async function readWorkbook(file) {
  const data = await file.arrayBuffer();
  return XLSX.read(data, { type: "array" });
}

async function readOptionalMasterRows(file, expectedHeaders) {
  if (!file) {
    return [];
  }
  const rows = await readDelimitedFile(file);
  return normalizeRowsToHeaders(rows, expectedHeaders);
}

async function readDelimitedFile(file) {
  const text = await file.text();
  return parseDelimited(text);
}

function parseDelimited(text) {
  const lines = String(text || "").replace(/^\uFEFF/, "").split(/\r?\n/);
  const firstLine = lines.find((line) => line.trim().length > 0) || "";
  const delimiter = firstLine.includes("\t") ? "\t" : ",";
  return parseSeparated(text, delimiter);
}

function parseSeparated(text, delimiter) {
  const rows = [];
  let current = "";
  let row = [];
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch === delimiter && !inQuotes) {
      row.push(current);
      current = "";
      continue;
    }

    if ((ch === "\n" || ch === "\r") && !inQuotes) {
      if (ch === "\r" && next === "\n") {
        i += 1;
      }
      row.push(current);
      rows.push(row);
      row = [];
      current = "";
      continue;
    }

    current += ch;
  }

  if (current.length > 0 || row.length > 0) {
    row.push(current);
    rows.push(row);
  }

  if (rows.length > 0 && rows[0][0]?.charCodeAt(0) === 0xfeff) {
    rows[0][0] = rows[0][0].slice(1);
  }

  return rows.filter((r) => r.some((v) => String(v || "").trim() !== ""));
}

function normalizeRowsToHeaders(rawRows, expectedHeaders) {
  if (!rawRows || rawRows.length === 0) {
    return [];
  }

  const sourceHeaders = rawRows[0].map((h) => String(h || "").trim());
  const rows = [];

  for (let rowIndex = 1; rowIndex < rawRows.length; rowIndex += 1) {
    const sourceRow = rawRows[rowIndex];
    const row = createEmptyRow(expectedHeaders);

    sourceHeaders.forEach((header, index) => {
      if (expectedHeaders.includes(header)) {
        row[header] = sourceRow[index] ?? "";
      }
    });

    if (expectedHeaders.includes("企業名") && !row["企業名"] && sourceHeaders.length >= 3) {
      row["企業名"] = sourceRow[2] ?? "";
    }
    if (expectedHeaders.includes("氏名") && !row["氏名"] && sourceHeaders.length >= 5) {
      row["氏名"] = sourceRow[4] ?? sourceRow[0] ?? "";
    }

    if (Object.values(row).every((v) => String(v || "").trim() === "")) {
      continue;
    }

    rows.push(row);
  }

  return rows;
}

function escapeCsvCell(value) {
  const text = String(value ?? "");
  if (/[",\n\r]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function toCsv(rows, headers) {
  const lines = [];
  lines.push(headers.map(escapeCsvCell).join(","));
  for (const row of rows) {
    lines.push(headers.map((header) => escapeCsvCell(row[header] ?? "")).join(","));
  }
  return lines.join("\r\n");
}
