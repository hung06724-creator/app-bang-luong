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

function mapEmployee(row, currentBranch, rowNumber, sheet) {
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
    deductionAdvance: toNumber(sheet?.['R' + rowNumber]?.v),
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

    employees.push(mapEmployee(row, currentBranch, excelRowNumber, sheet));
  });

  FILE_STATE.workbookName = originalName;
  FILE_STATE.sheetName = preferredSheetName;
  FILE_STATE.headers = headers;
  FILE_STATE.rawRows = rows;
  FILE_STATE.employees = employees;
  FILE_STATE.branches = branches;
}

function renderLayout(content, title = "Bảng lương GEOL") {
  return `<!doctype html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(title)}</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font-family: 'Inter', Arial, sans-serif; background: #f0f4f8; color: #1e293b; line-height: 1.5; }
    
    .header { background: #ffffff; padding: 12px 24px; color: #1e3a8a; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-bottom: 3px solid #ca8a04; position: sticky; top: 0; z-index: 50; }
    .header-inner { max-width: 1180px; margin: 0 auto; display: flex; align-items: center; justify-content: space-between; }
    .header-left { display: flex; align-items: center; gap: 12px; font-weight: 700; font-size: 20px; letter-spacing: 0.5px; }
    .header-logo { max-height: 48px; border-radius: 8px; }
    .header-link { color: #1e3a8a; text-decoration: none; font-weight: 600; font-size: 15px; padding: 8px 16px; border-radius: 8px; background: #f8fafc; transition: all 0.2s; border: 1px solid #e2e8f0; }
    .header-link:hover { background: #eff6ff; color: #1d4ed8; border-color: #bfdbfe; }

    .wrap { max-width: 1180px; margin: 32px auto; padding: 0 24px; }
    .card { background: #fff; border-radius: 16px; padding: 32px; margin-bottom: 24px; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04); border: 1px solid #e2e8f0; }
    
    .title { font-size: 24px; font-weight: 700; margin: 0 0 8px; color: #1e3a8a; }
    .muted { color: #64748b; font-size: 15px; }
    .row { display: flex; gap: 12px; flex-wrap: wrap; }
    .grid { display: grid; gap: 16px; }
    .grid-2 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    .grid-3 { grid-template-columns: repeat(3, minmax(0, 1fr)); }
    
    .stat { background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%); border-radius: 16px; padding: 20px; border: 1px solid #e2e8f0; border-left: 4px solid #ca8a04; box-shadow: 0 2px 8px rgba(0,0,0,0.02); }
    .stat .label { font-size: 13px; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
    .stat .value { font-size: 28px; font-weight: 700; margin-top: 8px; color: #1e3a8a; }
    
    .btn { display: inline-flex; align-items: center; justify-content: center; text-decoration: none; border: none; background: #ca8a04; color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-weight: 600; font-size: 15px; transition: all 0.2s ease; }
    .btn:hover { background: #b47a03; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(202, 138, 4, 0.2); }
    .btn-blue { background: #1e3a8a; }
    .btn-blue:hover { background: #172554; box-shadow: 0 4px 12px rgba(30, 58, 138, 0.2); }
    .btn-outline { background: white; color: #1e3a8a; border: 1px solid #cbd5e1; }
    .btn-outline:hover { background: #f8fafc; color: #1d4ed8; border-color: #94a3b8; }
    
    input[type=file]::file-selector-button { background: #e0e7ff; background: #1e3a8a; color: white; border: none; padding: 8px 16px; border-radius: 6px; font-weight: 500; cursor: pointer; margin-right: 12px; transition: background 0.2s; }
    input[type=file]::file-selector-button:hover { background: #172554; }
    
    input[type=text], select, input[type=file] {
      width: 100%;
      padding: 12px 16px;
      border-radius: 8px;
      border: 1px solid #cbd5e1;
      background: #f8fafc;
      font-size: 15px;
      font-family: inherit;
      transition: all 0.2s;
    }
    input[type=text]:focus, select:focus { outline: none; border-color: #1e3a8a; background: #fff; box-shadow: 0 0 0 3px rgba(30, 58, 138, 0.1); }
    
    table { width: 100%; border-collapse: separate; border-spacing: 0; }
    th, td { padding: 14px 16px; text-align: left; vertical-align: middle; border-bottom: 1px solid #f1f5f9; }
    th { background: #f8fafc; color: #475569; font-weight: 600; text-transform: uppercase; font-size: 13px; border-bottom: 2px solid #e2e8f0; border-top: 1px solid #e2e8f0; }
    tr:last-child td { border-bottom: none; }
    tr:hover td { background: #f8fafc; }
    
    .pill { display: inline-block; background: #e0e7ff; color: #3730a3; padding: 6px 14px; border-radius: 999px; font-size: 13px; font-weight: 600; margin-right: 8px; margin-bottom: 8px; border: 1px solid #c7d2fe; }
    
    @media (max-width: 800px) {
      .grid-2, .grid-3 { grid-template-columns: 1fr; }
      .wrap { padding: 16px; margin: 0 auto; }
      .card { padding: 20px; }
      .header-inner { flex-direction: row; gap: 12px; }
      .header-left span { display: none; }
    }

    @media print {
      .no-print, .header { display: none !important; }
      body { background: white; }
      .card { box-shadow: none; padding: 0; margin: 0; border: none; }
      .wrap { max-width: 100%; padding: 0; margin: 0; }
    }
  </style>
</head>
<body>
  <div class="header no-print">
    <div class="header-inner">
      <div class="header-left">
        <img class="header-logo" src="/logo.png" alt="GEOL Logo" onerror="this.style.display='none'" />
        <span>Hệ Thống Bảng Lương GEOL</span>
      </div>
      <div>
        <a class="header-link" href="/">Màn hình chính</a>
      </div>
    </div>
  </div>
  <div class="wrap">${content}</div>
</body>
</html>`;
}

function renderHomePage() {
  const html = `
    <div class="card" style="display: flex; gap: 32px; align-items: center; flex-wrap: wrap;">
      <div style="flex: 1; min-width: 300px;">
        <h1 class="title" style="font-size: 32px;">Chào mừng đến hệ thống Payroll</h1>
        <p class="muted" style="margin-bottom: 24px;">Upload tệp tin Excel (.xlsx, .xls) của bạn để phần mềm tiến hành số hóa, định dạng chuẩn và tạo thành các phiếu lương có thể tải xuống dễ dàng.</p>
        
        <form class="no-print" action="/upload" method="POST" enctype="multipart/form-data" style="background: #f8fafc; border: 2px dashed #cbd5e1; padding: 24px; border-radius: 12px; text-align: center;">
          <input type="file" name="excelFile" accept=".xlsx,.xls,.csv" required style="border: none; background: transparent; padding: 0; box-shadow: none;" />
          <div style="height:16px"></div>
          <button class="btn btn-blue" type="submit" style="width: 100%;">Tải lên & Xử lý</button>
        </form>
      </div>
      
      ${FILE_STATE.employees.length ? `
      <div style="flex: 1; min-width: 300px; background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 16px; padding: 24px;">
        <h2 style="margin-top:0; color: #1e3a8a;">Dữ liệu hiện tại</h2>
        <div style="margin-bottom: 16px;">
          <div class="muted">Tệp đang xử lý</div>
          <div style="font-weight: 600; color: #0f172a; font-size: 16px;">${escapeHtml(FILE_STATE.workbookName)}</div>
        </div>
        <div style="margin-bottom: 24px;">
          <div class="muted">Phân vùng Sheet</div>
          <div style="font-weight: 600; color: #0f172a; font-size: 16px;">${escapeHtml(FILE_STATE.sheetName)}</div>
        </div>
        <div class="row no-print">
          <a class="btn" href="/employees" style="width: 100%; text-align: center;">Vào kho lưu trữ nhân viên</a>
        </div>
      </div>
      ` : `
      <div style="flex: 1; min-width: 300px; padding: 24px; text-align: center; border: 1px dashed #e2e8f0; border-radius: 16px;">
        <img src="/logo.png" style="max-height: 100px; opacity: 0.1; margin-bottom: 16px;" onerror="this.style.display='none'">
        <p class="muted">Chưa có dữ liệu nào được báo cáo.<br>Hãy tải file gốc của bạn lên.</p>
      </div>
      `}
    </div>

    <div class="card">
      <h2 style="margin-top:0; color: #1e3a8a;">Tiêu chuẩn các trường tự động nhận diện</h2>
      <p class="muted" style="margin-bottom: 20px;">
        Phần mềm tự động phát hiện hàng tiêu đề từ dòng số 3 và số 4. Tự động nhận biết các CƠ SỞ theo cột STT. Dưới đây là các cột dữ liệu quan trọng đang được ánh xạ xử lý:
      </p>
      <div>
        <span class="pill">Tên</span>
        <span class="pill">Lương cơ bản</span>
        <span class="pill">Phụ cấp</span>
        <span class="pill">Thu nhập khác</span>
        <span class="pill">Thưởng & Sinh nhật</span>
        <span class="pill">Lương tăng ca</span>
        <span class="pill">Ngày công nghỉ</span>
        <span class="pill">Trừ tạm ứng</span>
        <span class="pill">Trừ đã TT</span>
        <span class="pill">Bảo hiểm (BHXH, BHYT, BHTN)</span>
        <span class="pill">Thuế TNCN</span>
        <span class="pill" style="background: #1e3a8a; color: white;">Thực lĩnh</span>
      </div>
    </div>
  `;

  return renderLayout(html, "Ứng dụng bảng lương");
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
    <div class="row no-print" style="margin-bottom: 24px; display: flex; justify-content: space-between; align-items: center;">
      <div>
        <h1 class="title" style="margin: 0;">Quản lý dữ liệu nhân sự</h1>
        <p class="muted" style="margin: 4px 0 0 0;">Tệp phân tích: <strong>${escapeHtml(FILE_STATE.workbookName)}</strong></p>
      </div>
      <div>
        <a class="btn btn-outline" href="/">Xử lý file mới</a>
      </div>
    </div>

    <div class="card no-print">
      <form method="GET" action="/employees" class="grid grid-2">
        <div>
          <label style="display: block; font-size: 13px; font-weight: 600; color: #475569; margin-bottom: 6px;">TÌM KIẾM NHÂN VIÊN</label>
          <input
            type="text"
            name="q"
            value="${escapeHtml(q)}"
            placeholder="Nhập tên hoặc mã nhân viên..."
          />
        </div>

        <div>
          <label style="display: block; font-size: 13px; font-weight: 600; color: #475569; margin-bottom: 6px;">LỌC THEO CƠ SỞ</label>
          <div style="display: flex; gap: 12px;">
            <select name="branch" onchange="this.form.submit()" style="flex: 1;">
              <option value="">Tất cả các cơ sở</option>
              ${FILE_STATE.branches
                .map((item) => {
                  const selected = item === branch ? "selected" : "";
                  return `<option value="${escapeHtml(item)}" ${selected}>${escapeHtml(item)}</option>`;
                })
                .join("")}
            </select>
            <button class="btn btn-blue" type="submit" style="white-space: nowrap;">Lọc</button>
            <a class="btn btn-outline" href="/employees" title="Xóa bộ lọc" style="padding: 10px; display: flex; align-items: center; justify-content: center;">
              <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
            </a>
          </div>
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

    <div id="slip-export-content" style="background: white; padding: 12px 24px; border-radius: 12px; margin-bottom: 24px; font-size: 14px; font-family: 'Times New Roman', Times, serif, 'Inter', sans-serif;">
      <div style="text-align: center; margin-bottom: 12px; border-bottom: 2px solid #1e3a8a; padding-bottom: 10px;">
        <img src="/logo.png" alt="GEOL Logo" style="max-height: 56px;" onerror="this.style.display='none'" />
        <h1 class="title" style="margin-top: 8px; color: #1e3a8a; font-size: 20px; text-transform: uppercase;">PHIẾU LƯƠNG NHÂN VIÊN</h1>
      </div>

      <div style="margin-bottom: 12px;">
        <table style="width: 100%; border: none;">
          <tr>
            <td style="border: none; padding: 2px 0; width: 50%;"><strong>Họ tên:</strong> ${escapeHtml(employee.name)}</td>
            <td style="border: none; padding: 2px 0; width: 50%;"><strong>Chức danh:</strong> ${escapeHtml(employee.title)}</td>
          </tr>
          <tr>
            <td style="border: none; padding: 2px 0;"><strong>Mã nhân viên:</strong> ${escapeHtml(employee.employeeId)}</td>
            <td style="border: none; padding: 2px 0;"><strong>Cơ sở làm việc:</strong> ${escapeHtml(employee.branch)}</td>
          </tr>
        </table>
      </div>

      <h3 style="border-bottom: 1px solid #cbd5e1; padding-bottom: 4px; margin-bottom: 10px; color: #1e3a8a; font-size: 15px; font-weight: bold; text-transform: uppercase;">I. CHI TIẾT THU NHẬP VÀ KHẤU TRỪ</h3>
      <table border="1" style="width: 100%; border-collapse: collapse; border-color: #94a3b8; margin-bottom: ${employee.otherIncome > 0 ? '16px' : '30px'};">
        <tbody>
          <tr><td style="padding: 6px; width: 60%;"><strong>Lương</strong></td><td style="padding: 6px; text-align: right;">${money(employee.salary)}</td></tr>
          <tr><td style="padding: 6px;"><strong>Phụ cấp</strong></td><td style="padding: 6px; text-align: right;">${money(employee.allowance)}</td></tr>
          <tr><td style="padding: 6px;"><strong>Thu nhập khác</strong></td><td style="padding: 6px; text-align: right;">${money(employee.otherIncome)}</td></tr>
          <tr><td style="padding: 6px;"><strong>Thưởng/Sinh nhật</strong></td><td style="padding: 6px; text-align: right;">${money(employee.birthdayBonus)}</td></tr>
          <tr><td style="padding: 6px;"><strong>Tăng ca</strong></td><td style="padding: 6px; text-align: right;">${money(employee.overtime)}</td></tr>
          <tr><td style="padding: 6px;"><strong>Ngày công nghỉ</strong></td><td style="padding: 6px; text-align: right;">${escapeHtml(employee.daysOff)}</td></tr>
          <tr><td style="padding: 6px;"><strong>Trừ tháng trước chuyển sang</strong></td><td style="padding: 6px; text-align: right;">${money(employee.previousMonthCarry)}</td></tr>
          <tr><td style="padding: 6px;"><strong>Trừ tạm ứng</strong></td><td style="padding: 6px; text-align: right;">${money(employee.deductionAdvance)}</td></tr>
          <tr><td style="padding: 6px;"><strong>Trừ đã thanh toán</strong></td><td style="padding: 6px; text-align: right;">${money(employee.deductionPaid)}</td></tr>
          <tr><td style="padding: 6px;"><strong>Bảo hiểm Xã hội (BHXH, BHYT, BHTN)</strong></td><td style="padding: 6px; text-align: right;">${money(employee.socialInsurance + employee.healthInsurance + employee.unemploymentInsurance)}</td></tr>
          <tr><td style="padding: 6px;"><strong>Thuế Thu nhập Cá nhân (TNCN)</strong></td><td style="padding: 6px; text-align: right;">${money(employee.pitTax)}</td></tr>
          <tr>
            <td style="padding: 8px; background: #e0e7ff; color: #1e3a8a; font-size: 15px;"><strong>THỰC LĨNH</strong></td>
            <td style="padding: 8px; text-align: right; background: #e0e7ff; color: #1e3a8a; font-size: 16px; font-weight: bold;">${money(employee.netIncome)}</td>
          </tr>
        </tbody>
      </table>

      ${employee.otherIncome > 0 ? `
      <h3 style="border-bottom: 1px solid #cbd5e1; padding-bottom: 4px; margin-bottom: 10px; color: #1e3a8a; font-size: 15px; font-weight: bold; text-transform: uppercase;">II. CHI TIẾT THU NHẬP KHÁC</h3>
      <table border="1" style="width: 100%; border-collapse: collapse; border-color: #94a3b8;">
        <tbody>
          <tr><td style="padding: 6px; width: 60%;">Hoa hồng tuyển sinh</td><td style="padding: 6px; text-align: right;">${money(employee.otherIncomeBreakdown.admissionsCommission)}</td></tr>
          <tr><td style="padding: 6px;">Thưởng tết</td><td style="padding: 6px; text-align: right;">${money(employee.otherIncomeBreakdown.tetBonus)}</td></tr>
          <tr><td style="padding: 6px;">Thưởng chức vụ</td><td style="padding: 6px; text-align: right;">${money(employee.otherIncomeBreakdown.positionBonus)}</td></tr>
          <tr><td style="padding: 6px;">Học viên bay</td><td style="padding: 6px; text-align: right;">${money(employee.otherIncomeBreakdown.pilotStudent)}</td></tr>
          <tr><td style="padding: 6px;">Lương dạy online/trực page</td><td style="padding: 6px; text-align: right;">${money(employee.otherIncomeBreakdown.onlineTeaching)}</td></tr>
        </tbody>
      </table>
      ` : ''}
      
      <div style="margin-top: ${employee.otherIncome > 0 ? '16px' : '0'}; text-align: right; font-style: italic; color: #475569;">
        Hà Nội, Ngày ${new Date().getDate()} tháng ${new Date().getMonth() + 1} năm ${new Date().getFullYear()}
      </div>
    </div>

    <script>
      function exportPDF() {
        const element = document.getElementById('slip-export-content');
        const opt = {
          margin:       0.25,
          filename:     'Phieu_Luong_${employeeNameSafe}.pdf',
          image:        { type: 'jpeg', quality: 0.98 },
          html2canvas:  { scale: 2 },
          jsPDF:        { unit: 'in', format: 'a4', orientation: 'portrait' }
        };
        html2pdf().set(opt).from(element).save();
      }

      function exportWord() {
        const element = document.getElementById('slip-export-content');
        let htmlContent = element.innerHTML;
        
        // Fix image path for Word (convert relative /logo.png to absolute URL so Word can handle it properly online)
        const absoluteUrl = window.location.origin + '/logo.png';
        htmlContent = htmlContent.replace(/src="\\/logo\\.png"/g, 'src="' + absoluteUrl + '"');

        // Prepend word schema
        const preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Phiếu Lương</title></head><body>";
        const postHtml = "</body></html>";
        const html = preHtml + htmlContent + postHtml;

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
