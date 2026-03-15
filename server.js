const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
const os = require("os");

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({ dest: os.tmpdir() });

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static("public"));

const FILE_STATE = {
  workbookName: "",
  sheetName: "",
  headers: [],
  rawRows: [],
  employees: [],
  branches: [],
};

const COLUMN_MAP = {
  id: "Mã nhân viên",
  name: "Tên nhân viên",
  title: "Chức danh",
  salary: "Lương cơ bản",
  allowance: "Phụ cấp thuộc quỹ lương",
  otherIncomeParts: [
    "Thu nhập khác ( Hoa hồng) - Hoa hồng tuyển sinh",
    "Thu nhập khác ( Hoa hồng) - Thưởng tết",
    "Thu nhập khác ( Hoa hồng) - Thưởng chức vụ (GĐ2)",
    "Thu nhập khác ( Hoa hồng) - Học viên bay",
    "Thu nhập khác ( Hoa hồng) - Lương dạy online/trực page",
  ],
  birthdayBonus: "Thưởng/Sinh nhật",
  overtime: "Tăng ca",
  daysOff: "Ngày công nghỉ",
  previousMonthCarry: "Trừ tháng trước chuyển sang",
  deductionAdvance: "Trừ - tạm ứng",
  deductionPaid: "Trừ - đã tt",
  insuranceBaseSalary: "Lương đóng bảo hiểm",
  socialInsurance: "Các khoản khấu trừ - BHXH",
  healthInsurance: "Các khoản khấu trừ - BHYT",
  unemploymentInsurance: "Các khoản khấu trừ - BHTN",
  unionFee: "Các khoản khấu trừ - KPCĐ",
  pitTax: "Các khoản khấu trừ - Thuế TNCN",
  totalDeduction: "Các khoản khấu trừ - Cộng",
  familyDeduction: "Giảm trừ gia cảnh",
  taxableIncomeBeforeDeduction: "Tổng thu nhập chịu thuế TNCN",
  taxableIncome: "Thu nhập tính thuế TNCN",
  netIncome: "Số tiền còn được lĩnh",
};

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function toNumber(value) {
  if (value == null || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;

  const raw = String(value).trim();
  if (!raw) return 0;

  const cleaned = raw
    .replace(/\s+/g, "")
    .replace(/\.(?=\d{3}(\D|$))/g, "")
    .replace(/,/g, ".")
    .replace(/[^0-9.-]/g, "");

  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function money(value) {
  return new Intl.NumberFormat("vi-VN", {
    style: "currency",
    currency: "VND",
    maximumFractionDigits: 0,
  }).format(toNumber(value));
}

function buildHeadersFromRow3And4(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const headerTopRow = 2; // Excel row 3
  const headerSubRow = 3; // Excel row 4

  const headers = [];
  let lastTopHeader = "";

  for (let c = range.s.c; c <= range.e.c; c++) {
    const topCell = sheet[XLSX.utils.encode_cell({ r: headerTopRow, c })];
    const subCell = sheet[XLSX.utils.encode_cell({ r: headerSubRow, c })];

    const top = String(topCell?.v || "")
      .replace(/\n/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    const sub = String(subCell?.v || "")
      .replace(/\n/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    if (top) lastTopHeader = top;

    let finalHeader = "";
    if (top && sub) {
      finalHeader = `${top} - ${sub}`;
    } else if (!top && sub) {
      finalHeader = `${lastTopHeader} - ${sub}`;
    } else if (top && !sub) {
      finalHeader = top;
    } else {
      finalHeader = `COL_${c}`;
    }

    headers.push(finalHeader.replace(/\s+/g, " ").trim());
  }

  return headers;
}

function isBranchRow(row) {
  const firstCell = String(row["STT"] || "").trim();
  return normalizeText(firstCell).startsWith("co so ");
}

function getBranchName(row) {
  return String(row["STT"] || "").trim();
}

function isLikelyEmployeeRow(row) {
  if (isBranchRow(row)) return false;

  const name = String(row[COLUMN_MAP.name] || "").trim();
  const employeeId = String(row[COLUMN_MAP.id] || "").trim();
  const title = String(row[COLUMN_MAP.title] || "").trim();
  const salary = toNumber(row[COLUMN_MAP.salary]);
  const allowance = toNumber(row[COLUMN_MAP.allowance]);
  const netIncome = toNumber(row[COLUMN_MAP.netIncome]);

  if (!name) return false;
  if (normalizeText(name).startsWith("tong")) return false;

  return Boolean(employeeId || title || salary || allowance || netIncome);
}

function sumOtherIncome(row) {
  return COLUMN_MAP.otherIncomeParts.reduce((sum, key) => {
    return sum + toNumber(row[key]);
  }, 0);
}

function mapEmployee(row, currentBranch, rowNumber) {
  return {
    excelRowNumber: rowNumber,
    branch: currentBranch,
    employeeId: String(row[COLUMN_MAP.id] || "").trim(),
    name: String(row[COLUMN_MAP.name] || "").trim(),
    title: String(row[COLUMN_MAP.title] || "").trim(),
    salary: toNumber(row[COLUMN_MAP.salary]),
    allowance: toNumber(row[COLUMN_MAP.allowance]),
    otherIncome: sumOtherIncome(row),
    otherIncomeBreakdown: {
      admissionsCommission: toNumber(row[COLUMN_MAP.otherIncomeParts[0]]),
      tetBonus: toNumber(row[COLUMN_MAP.otherIncomeParts[1]]),
      positionBonus: toNumber(row[COLUMN_MAP.otherIncomeParts[2]]),
      pilotStudent: toNumber(row[COLUMN_MAP.otherIncomeParts[3]]),
      onlineTeaching: toNumber(row[COLUMN_MAP.otherIncomeParts[4]]),
    },
    birthdayBonus: toNumber(row[COLUMN_MAP.birthdayBonus]),
    overtime: toNumber(row[COLUMN_MAP.overtime]),
    daysOff: row[COLUMN_MAP.daysOff] ?? "",
    previousMonthCarry: toNumber(row[COLUMN_MAP.previousMonthCarry]),
    deductionAdvance: toNumber(row[COLUMN_MAP.deductionAdvance]),
    deductionPaid: toNumber(row[COLUMN_MAP.deductionPaid]),
    insuranceBaseSalary: toNumber(row[COLUMN_MAP.insuranceBaseSalary]),
    socialInsurance: toNumber(row[COLUMN_MAP.socialInsurance]),
    healthInsurance: toNumber(row[COLUMN_MAP.healthInsurance]),
    unemploymentInsurance: toNumber(row[COLUMN_MAP.unemploymentInsurance]),
    unionFee: toNumber(row[COLUMN_MAP.unionFee]),
    pitTax: toNumber(row[COLUMN_MAP.pitTax]),
    totalDeduction: toNumber(row[COLUMN_MAP.totalDeduction]),
    familyDeduction: toNumber(row[COLUMN_MAP.familyDeduction]),
    taxableIncomeBeforeDeduction: toNumber(row[COLUMN_MAP.taxableIncomeBeforeDeduction]),
    taxableIncome: toNumber(row[COLUMN_MAP.taxableIncome]),
    netIncome: toNumber(row[COLUMN_MAP.netIncome]),
  };
}

function parsePayrollFile(filePath, originalName) {
  const workbook = XLSX.readFile(filePath, {
    cellFormula: false,
    cellHTML: false,
    cellNF: false,
    cellText: false,
  });

  const preferredSheetName =
    workbook.SheetNames.find((name) =>
      normalizeText(name).includes("bang luong")
    ) || workbook.SheetNames[0];

  const sheet = workbook.Sheets[preferredSheetName];
  const headers = buildHeadersFromRow3And4(sheet);

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: headers,
    range: 4, // bắt đầu từ Excel row 5
    defval: "",
    blankrows: false,
  });

  let currentBranch = "";
  const employees = [];
  const branches = [];

  rows.forEach((row, index) => {
    const excelRowNumber = index + 5;

    if (isBranchRow(row)) {
      currentBranch = getBranchName(row);
      if (currentBranch && !branches.includes(currentBranch)) {
        branches.push(currentBranch);
      }
      return;
    }

    if (!isLikelyEmployeeRow(row)) return;

    employees.push(mapEmployee(row, currentBranch, excelRowNumber));
  });

  FILE_STATE.workbookName = originalName;
  FILE_STATE.sheetName = preferredSheetName;
  FILE_STATE.headers = headers;
  FILE_STATE.rawRows = rows;
  FILE_STATE.employees = employees;
  FILE_STATE.branches = branches;
}

function renderLayout(content, title = "Bảng lương local") {
  return `<!doctype html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(title)}</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font-family: Arial, sans-serif; background: #f8fafc; color: #0f172a; }
    .wrap { max-width: 1180px; margin: 0 auto; padding: 24px; }
    .card { background: #fff; border-radius: 18px; padding: 20px; margin-bottom: 18px; box-shadow: 0 1px 10px rgba(15, 23, 42, 0.08); }
    .title { font-size: 28px; font-weight: 700; margin: 0 0 10px; }
    .muted { color: #64748b; }
    .row { display: flex; gap: 12px; flex-wrap: wrap; }
    .grid { display: grid; gap: 14px; }
    .grid-2 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    .grid-3 { grid-template-columns: repeat(3, minmax(0, 1fr)); }
    .stat { background: #f8fafc; border-radius: 14px; padding: 16px; }
    .stat .label { font-size: 13px; color: #64748b; }
    .stat .value { font-size: 24px; font-weight: 700; margin-top: 6px; }
    .btn { display: inline-block; text-decoration: none; border: none; background: #0f172a; color: white; padding: 10px 14px; border-radius: 12px; cursor: pointer; }
    .btn-outline { background: white; color: #0f172a; border: 1px solid #cbd5e1; }
    input[type=file], input[type=text], select {
      width: 100%;
      padding: 11px 12px;
      border-radius: 12px;
      border: 1px solid #cbd5e1;
    }
    table { width: 100%; border-collapse: collapse; }
    th, td {
      padding: 12px;
      border-bottom: 1px solid #e2e8f0;
      text-align: left;
      vertical-align: top;
    }
    th { background: #f8fafc; }
    .pill {
      display: inline-block;
      background: #e2e8f0;
      color: #0f172a;
      padding: 6px 10px;
      border-radius: 999px;
      font-size: 13px;
      margin-right: 8px;
      margin-bottom: 8px;
    }
    .pay-grid {
      display: grid;
      gap: 12px;
      grid-template-columns: repeat(2, minmax(0, 1fr));
    }
    .pay-item {
      background: #f8fafc;
      border-radius: 14px;
      padding: 14px;
    }
    .pay-item .k { color: #64748b; font-size: 13px; }
    .pay-item .v { font-weight: 700; margin-top: 6px; }
    .pay-item.highlight { background: #0f172a; color: white; }
    .pay-item.highlight .k { color: #cbd5e1; }
    .small { font-size: 13px; }

    @media (max-width: 800px) {
      .grid-2, .grid-3, .pay-grid { grid-template-columns: 1fr; }
    }

    @media print {
      .no-print { display: none !important; }
      body { background: white; }
      .card { box-shadow: none; padding: 0; }
      .wrap { max-width: 100%; padding: 0; }
    }
  </style>
</head>
<body>
  <div class="wrap">${content}</div>
</body>
</html>`;
}

function renderHomePage() {
  const html = `
    <div class="card">
      <h1 class="title">App bảng lương local - Node.js</h1>
      <p class="muted">Upload file Excel để đọc bảng lương, tự nhận diện cơ sở và tạo phiếu lương cho từng nhân viên.</p>

      <form class="no-print" action="/upload" method="POST" enctype="multipart/form-data">
        <input type="file" name="excelFile" accept=".xlsx,.xls,.csv" required />
        <div style="height:12px"></div>
        <button class="btn" type="submit">Tải file lên</button>
      </form>
    </div>

    <div class="card">
      <h2 style="margin-top:0;">Logic đã áp dụng cho file này</h2>
      <p class="muted small">
        Đọc header từ dòng 3 và 4, sau đó tự ghép tên cột.
        Các dòng bắt đầu bằng “CƠ SỞ ...” ở cột STT sẽ được coi là dòng cơ sở và không tính là nhân viên.
      </p>
      <div>
        <span class="pill">Tên</span>
        <span class="pill">Lương</span>
        <span class="pill">Phụ cấp</span>
        <span class="pill">Thu nhập khác</span>
        <span class="pill">Thưởng/Sinh nhật</span>
        <span class="pill">Tăng ca</span>
        <span class="pill">Ngày công nghỉ</span>
        <span class="pill">Trừ tạm ứng</span>
        <span class="pill">Trừ đã TT</span>
        <span class="pill">BHXH, BHYT, BHTN</span>
        <span class="pill">Thuế TNCN</span>
        <span class="pill">Thực lĩnh</span>
      </div>
    </div>

    ${FILE_STATE.employees.length ? `
      <div class="card">
        <h2 style="margin-top:0;">Dữ liệu đã nạp</h2>
        <p class="muted">
          File: <strong>${escapeHtml(FILE_STATE.workbookName)}</strong>
          |
          Sheet: <strong>${escapeHtml(FILE_STATE.sheetName)}</strong>
        </p>
        <div class="row no-print">
          <a class="btn" href="/employees">Xem danh sách nhân viên</a>
        </div>
      </div>
    ` : ""}
  `;

  return renderLayout(html, "App bảng lương local");
}

function renderEmployeesPage(req) {
  const q = String(req.query.q || "").trim();
  const branch = String(req.query.branch || "").trim();
  const qNorm = normalizeText(q);
  const branchNorm = normalizeText(branch);

  const employees = FILE_STATE.employees.filter((employee) => {
    const matchName =
      !qNorm ||
      normalizeText(employee.name).includes(qNorm) ||
      normalizeText(employee.employeeId).includes(qNorm);

    const matchBranch =
      !branchNorm || normalizeText(employee.branch) === branchNorm;

    return matchName && matchBranch;
  });

  const totalSalary = employees.reduce((sum, item) => sum + item.salary, 0);
  const totalNetIncome = employees.reduce((sum, item) => sum + item.netIncome, 0);

  const html = `
    <div class="card no-print">
      <h1 class="title">Danh sách nhân viên</h1>
      <p class="muted">File: <strong>${escapeHtml(FILE_STATE.workbookName)}</strong></p>

      <form method="GET" action="/employees" class="grid grid-2">
        <div>
          <label class="small muted">Tìm theo tên hoặc mã nhân viên</label>
          <input
            type="text"
            name="q"
            value="${escapeHtml(q)}"
            placeholder="Ví dụ: Trần Thị Hiền hoặc NV00028"
          />
        </div>

        <div>
          <label class="small muted">Lọc theo cơ sở</label>
          <select name="branch" onchange="this.form.submit()">
            <option value="">Tất cả cơ sở</option>
            ${FILE_STATE.branches
              .map((item) => {
                const selected = item === branch ? "selected" : "";
                return `<option value="${escapeHtml(item)}" ${selected}>${escapeHtml(item)}</option>`;
              })
              .join("")}
          </select>
        </div>

        <div>
          <button class="btn" type="submit">Lọc dữ liệu</button>
          <a class="btn btn-outline" href="/employees" style="margin-left:8px;">Xóa lọc</a>
        </div>
      </form>
    </div>

    <div class="card">
      <div class="grid grid-3">
        <div class="stat">
          <div class="label">Số nhân viên</div>
          <div class="value">${employees.length}</div>
        </div>
        <div class="stat">
          <div class="label">Tổng lương</div>
          <div class="value">${money(totalSalary)}</div>
        </div>
        <div class="stat">
          <div class="label">Tổng thực lĩnh</div>
          <div class="value">${money(totalNetIncome)}</div>
        </div>
      </div>
    </div>

    <div class="card">
      <table>
        <thead>
          <tr>
            <th>Mã NV</th>
            <th>Tên nhân viên</th>
            <th>Chức danh</th>
            <th>Cơ sở</th>
            <th>Lương</th>
            <th>Thực lĩnh</th>
            <th class="no-print">Phiếu lương</th>
          </tr>
        </thead>
        <tbody>
          ${employees
            .map((employee) => {
              const id = encodeURIComponent(employee.employeeId || employee.name);
              return `
                <tr>
                  <td>${escapeHtml(employee.employeeId)}</td>
                  <td>${escapeHtml(employee.name)}</td>
                  <td>${escapeHtml(employee.title)}</td>
                  <td>${escapeHtml(employee.branch)}</td>
                  <td>${money(employee.salary)}</td>
                  <td>${money(employee.netIncome)}</td>
                  <td class="no-print">
                    <a class="btn btn-outline" href="/slip/${id}">Xem phiếu</a>
                  </td>
                </tr>
              `;
            })
            .join("")}
        </tbody>
      </table>
    </div>

    <div class="row no-print">
      <a class="btn btn-outline" href="/">Tải file khác</a>
    </div>
  `;

  return renderLayout(html, "Danh sách nhân viên");
}

function renderSlipPage(employee) {
  const employeeNameSafe = employee.name.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase();

  const html = `
    <div class="card no-print">
      <div class="row">
        <a class="btn btn-outline" href="/employees">Quay lại</a>
        <button class="btn" onclick="window.print()">In phiếu</button>
        <button class="btn" onclick="exportPDF()" style="background:#2563eb;">Tải PDF</button>
        <button class="btn" onclick="exportWord()" style="background:#0ea5e9;">Tải Word</button>
      </div>
    </div>

    <div id="slip-export-content" style="background: white; padding: 12px 24px; border-radius: 12px; margin-bottom: 24px; font-size: 14px;">
      <div style="text-align: center; margin-bottom: 12px; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px;">
        <img src="/logo.png" alt="GEOL Logo" style="max-height: 60px;" onerror="this.style.display='none'" />
        <h1 class="title" style="margin-top: 8px; color: #0f172a; font-size: 20px;">PHIẾU LƯƠNG NHÂN VIÊN</h1>
      </div>

      <div style="margin-bottom: 12px;">
        <table style="width: 100%; border: none;">
          <tr>
            <td style="border: none; padding: 2px 0;"><strong>Họ tên:</strong> ${escapeHtml(employee.name)}</td>
            <td style="border: none; padding: 2px 0;"><strong>Chức danh:</strong> ${escapeHtml(employee.title)}</td>
          </tr>
          <tr>
            <td style="border: none; padding: 2px 0;"><strong>Mã nhân viên:</strong> ${escapeHtml(employee.employeeId)}</td>
            <td style="border: none; padding: 2px 0;"><strong>Cơ sở làm việc:</strong> ${escapeHtml(employee.branch)}</td>
          </tr>
        </table>
      </div>

      <h3 style="border-bottom: 1px solid #cbd5e1; padding-bottom: 4px; margin-bottom: 10px; color: #334155; font-size: 16px;">I. CHI TIẾT THU NHẬP VÀ KHẤU TRỪ</h3>
      <table border="1" style="width: 100%; border-collapse: collapse; border-color: #cbd5e1; margin-bottom: ${employee.otherIncome > 0 ? '16px' : '40px'};">
        <tbody>
          <tr><td style="padding: 6px; background: #f8fafc; width: 60%;"><strong>Lương</strong></td><td style="padding: 6px; text-align: right;">${money(employee.salary)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;"><strong>Phụ cấp</strong></td><td style="padding: 6px; text-align: right;">${money(employee.allowance)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;"><strong>Thu nhập khác</strong></td><td style="padding: 6px; text-align: right;">${money(employee.otherIncome)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;"><strong>Thưởng/Sinh nhật</strong></td><td style="padding: 6px; text-align: right;">${money(employee.birthdayBonus)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;"><strong>Tăng ca</strong></td><td style="padding: 6px; text-align: right;">${money(employee.overtime)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;"><strong>Ngày công nghỉ</strong></td><td style="padding: 6px; text-align: right;">${escapeHtml(employee.daysOff)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;"><strong>Trừ tháng trước chuyển sang</strong></td><td style="padding: 6px; text-align: right;">${money(employee.previousMonthCarry)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;"><strong>Trừ tạm ứng</strong></td><td style="padding: 6px; text-align: right;">${money(employee.deductionAdvance)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;"><strong>Trừ đã thanh toán</strong></td><td style="padding: 6px; text-align: right;">${money(employee.deductionPaid)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;"><strong>Bảo hiểm Xã hội (BHXH, BHYT, BHTN)</strong></td><td style="padding: 6px; text-align: right;">${money(employee.socialInsurance + employee.healthInsurance + employee.unemploymentInsurance)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;"><strong>Thuế Thu nhập Cá nhân (TNCN)</strong></td><td style="padding: 6px; text-align: right;">${money(employee.pitTax)}</td></tr>
          <tr>
            <td style="padding: 8px; background: #0f172a; color: white; font-size: 16px;"><strong>THỰC LĨNH</strong></td>
            <td style="padding: 8px; text-align: right; background: #0f172a; color: white; font-size: 16px; font-weight: bold;">${money(employee.netIncome)}</td>
          </tr>
        </tbody>
      </table>

      ${employee.otherIncome > 0 ? `
      <h3 style="border-bottom: 1px solid #cbd5e1; padding-bottom: 4px; margin-bottom: 10px; color: #334155; font-size: 16px;">II. CHI TIẾT THU NHẬP KHÁC</h3>
      <table border="1" style="width: 100%; border-collapse: collapse; border-color: #cbd5e1;">
        <tbody>
          <tr><td style="padding: 6px; background: #f8fafc; width: 60%;">Hoa hồng tuyển sinh</td><td style="padding: 6px; text-align: right;">${money(employee.otherIncomeBreakdown.admissionsCommission)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;">Thưởng tết</td><td style="padding: 6px; text-align: right;">${money(employee.otherIncomeBreakdown.tetBonus)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;">Thưởng chức vụ</td><td style="padding: 6px; text-align: right;">${money(employee.otherIncomeBreakdown.positionBonus)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;">Học viên bay</td><td style="padding: 6px; text-align: right;">${money(employee.otherIncomeBreakdown.pilotStudent)}</td></tr>
          <tr><td style="padding: 6px; background: #f8fafc;">Lương dạy online/trực page</td><td style="padding: 6px; text-align: right;">${money(employee.otherIncomeBreakdown.onlineTeaching)}</td></tr>
        </tbody>
      </table>
      ` : ''}
      
      <div style="margin-top: ${employee.otherIncome > 0 ? '20px' : '0'}; text-align: right; font-style: italic; color: #64748b;">
        Ngày xuất phiếu: ${new Date().toLocaleDateString('vi-VN')}
      </div>
    </div>

    <script>
      function exportPDF() {
        const element = document.getElementById('slip-export-content');
        const opt = {
          margin:       0.3,
          filename:     'Phieu_Luong_${employeeNameSafe}.pdf',
          image:        { type: 'jpeg', quality: 0.98 },
          html2canvas:  { scale: 2 },
          jsPDF:        { unit: 'in', format: 'a4', orientation: 'portrait' }
        };
        html2pdf().set(opt).from(element).save();
      }

      function exportWord() {
        const element = document.getElementById('slip-export-content');
        // Prepend word schema
        const preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Phiếu Lương</title></head><body>";
        const postHtml = "</body></html>";
        const html = preHtml + element.innerHTML + postHtml;

        const blob = new Blob(['\\ufeff', html], {
            type: 'application/msword'
        });
        
        const downloadLink = document.createElement("a");
        document.body.appendChild(downloadLink);
        
        const url = URL.createObjectURL(blob);
        downloadLink.href = url;
        downloadLink.download = 'Phieu_Luong_${employeeNameSafe}.doc';
        downloadLink.click();
        
        URL.revokeObjectURL(url);
        document.body.removeChild(downloadLink);
      }
    </script>
  `;

  return renderLayout(html, `Phiếu lương - ${employee.name}`);
}

app.get("/", (req, res) => {
  res.send(renderHomePage());
});

app.post("/upload", upload.single("excelFile"), (req, res) => {
  if (!req.file) {
    return res
      .status(400)
      .send(
        renderLayout(
          `<div class="card"><h2>Lỗi</h2><p>Chưa có file Excel.</p><a class="btn" href="/">Quay lại</a></div>`,
          "Lỗi upload"
        )
      );
  }

  try {
    parsePayrollFile(req.file.path, req.file.originalname);
    fs.unlink(req.file.path, () => {});
    res.redirect("/employees");
  } catch (error) {
    fs.unlink(req.file.path, () => {});
    res.status(500).send(
      renderLayout(
        `
        <div class="card">
          <h2>Lỗi xử lý file</h2>
          <p>${escapeHtml(error.message)}</p>
          <a class="btn" href="/">Quay lại</a>
        </div>
        `,
        "Lỗi xử lý file"
      )
    );
  }
});

app.get("/employees", (req, res) => {
  if (!FILE_STATE.employees.length) return res.redirect("/");
  res.send(renderEmployeesPage(req));
});

app.get("/slip/:id", (req, res) => {
  if (!FILE_STATE.employees.length) return res.redirect("/");

  const id = decodeURIComponent(req.params.id || "");
  const employee =
    FILE_STATE.employees.find((item) => item.employeeId === id) ||
    FILE_STATE.employees.find((item) => item.name === id);

  if (!employee) {
    return res.status(404).send(
      renderLayout(
        `
        <div class="card">
          <h2>Không tìm thấy nhân viên</h2>
          <a class="btn" href="/employees">Quay lại danh sách</a>
        </div>
        `,
        "Không tìm thấy nhân viên"
      )
    );
  }

  res.send(renderSlipPage(employee));
});

if (process.env.NODE_ENV !== "production") {
  app.listen(PORT, () => {
    console.log(`Server is running at http://localhost:${PORT}`);
  });
}

module.exports = app;
