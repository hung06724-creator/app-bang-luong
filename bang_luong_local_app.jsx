const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 3000;

const upload = multer({ dest: path.join(__dirname, "uploads") });

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

const money = (value) => {
  const num = Number(value || 0);
  return new Intl.NumberFormat("vi-VN", {
    style: "currency",
    currency: "VND",
    maximumFractionDigits: 0,
  }).format(num);
};

const normalize = (s) =>
  String(s || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();

const FIELD_LABELS = {
  name: "Tên",
  salary: "Lương",
  allowance: "Phụ cấp",
  otherIncome: "Thu nhập khác",
  birthdayBonus: "Thưởng/Sinh nhật",
  overtime: "Tăng ca",
  daysOff: "Ngày công nghỉ",
  advanceDeduction: "Trừ tạm ứng",
  paidDeduction: "Trừ đã TT",
  socialInsurance: "BHXH",
  pitTax: "Thuế TNCN",
  netIncome: "Thực lĩnh",
};

const FIELD_HINTS = {
  name: ["ten", "ho ten", "nhan vien", "employee", "name"],
  salary: ["luong", "luong chinh", "luong cb", "salary"],
  allowance: ["phu cap", "allowance"],
  otherIncome: ["thu nhap khac", "hh tuyen sinh", "thuong tet", "thuong chuc vu", "hoc vien bay", "luong onl", "luong online"],
  birthdayBonus: ["thuong sinh nhat", "sinh nhat", "birthday"],
  overtime: ["tang ca", "overtime", "ot"],
  daysOff: ["ngay cong nghi", "nghi", "ngay nghi"],
  advanceDeduction: ["tru tam ung", "tam ung"],
  paidDeduction: ["tru da tt", "da tt"],
  socialInsurance: ["bhxh", "bao hiem xa hoi"],
  pitTax: ["thue tncn", "tncn", "pit"],
  netIncome: ["thuc linh", "net", "take home"],
};

function guessColumn(headers, field) {
  const hints = FIELD_HINTS[field] || [];
  const normalizedHeaders = headers.map((h) => ({ raw: h, norm: normalize(h) }));

  for (const hint of hints) {
    const exact = normalizedHeaders.find((h) => h.norm === normalize(hint));
    if (exact) return exact.raw;
  }

  for (const hint of hints) {
    const partial = normalizedHeaders.find((h) => h.norm.includes(normalize(hint)));
    if (partial) return partial.raw;
  }

  return "";
}

function toNumber(value) {
  if (value == null || value === "") return 0;
  if (typeof value === "number") return value;
  const cleaned = String(value)
    .replace(/\./g, "")
    .replace(/,/g, ".")
    .replace(/[^0-9.-]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function buildEmployee(row, mapping) {
  return {
    name: row[mapping.name] || "",
    salary: toNumber(row[mapping.salary]),
    allowance: toNumber(row[mapping.allowance]),
    otherIncome: toNumber(row[mapping.otherIncome]),
    birthdayBonus: toNumber(row[mapping.birthdayBonus]),
    overtime: toNumber(row[mapping.overtime]),
    daysOff: row[mapping.daysOff] || "",
    advanceDeduction: toNumber(row[mapping.advanceDeduction]),
    paidDeduction: toNumber(row[mapping.paidDeduction]),
    socialInsurance: toNumber(row[mapping.socialInsurance]),
    pitTax: toNumber(row[mapping.pitTax]),
    netIncome: toNumber(row[mapping.netIncome]),
  };
}

function renderLayout(content, title = "App bảng lương local") {
  return `<!doctype html>
  <html lang="vi">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${title}</title>
    <style>
      body { font-family: Arial, sans-serif; background: #f8fafc; margin: 0; color: #0f172a; }
      .wrap { max-width: 1100px; margin: 0 auto; padding: 24px; }
      .card { background: white; border-radius: 20px; padding: 20px; box-shadow: 0 1px 8px rgba(0,0,0,0.06); margin-bottom: 20px; }
      h1,h2,h3,p { margin-top: 0; }
      .btn { display:inline-block; background:#0f172a; color:white; border:none; padding:10px 14px; border-radius:12px; text-decoration:none; cursor:pointer; }
      .btn-outline { background:white; color:#0f172a; border:1px solid #cbd5e1; }
      .grid { display:grid; gap:16px; }
      .grid-2 { grid-template-columns: 1fr 1fr; }
      .grid-3 { grid-template-columns: repeat(3, 1fr); }
      .item { background:#f8fafc; border-radius:16px; padding:14px; }
      .label { font-size:13px; color:#64748b; }
      .value { font-weight:700; margin-top:6px; }
      table { width:100%; border-collapse: collapse; }
      th, td { border-bottom:1px solid #e2e8f0; text-align:left; padding:12px; }
      input[type=file], input[type=text], select { width:100%; padding:10px; border:1px solid #cbd5e1; border-radius:12px; box-sizing:border-box; }
      .actions { display:flex; gap:10px; flex-wrap:wrap; }
      @media (max-width: 768px) {
        .grid-2, .grid-3 { grid-template-columns: 1fr; }
      }
      @media print {
        .no-print { display:none !important; }
        body { background:white; }
        .card { box-shadow:none; padding:0; }
      }
    </style>
  </head>
  <body>
    <div class="wrap">${content}</div>
  </body>
  </html>`;
}

let payrollData = {
  headers: [],
  rows: [],
  mapping: null,
  employees: [],
};

app.get("/", (req, res) => {
  const html = `
    <div class="card">
      <h1>App bảng lương local - Node.js</h1>
      <p>Upload file Excel để đọc dữ liệu và tạo phiếu lương cho từng nhân viên.</p>
      <form class="no-print" action="/upload" method="POST" enctype="multipart/form-data">
        <input type="file" name="excelFile" accept=".xlsx,.xls,.csv" required />
        <br /><br />
        <button class="btn" type="submit">Tải file lên</button>
      </form>
    </div>

    <div class="card">
      <h3>Các trường đang hỗ trợ</h3>
      <p>${Object.values(FIELD_LABELS).join(" • ")}</p>
      <p>Nếu tên cột trong file khác nhau, hệ thống sẽ tự đoán gần đúng. Khi bạn gửi file thật, tôi có thể chỉnh lại mapping chuẩn 100%.</p>
    </div>

    ${payrollData.employees.length ? `
      <div class="card">
        <div class="actions no-print">
          <a class="btn" href="/employees">Xem danh sách nhân viên</a>
        </div>
        <p>Đã đọc <strong>${payrollData.employees.length}</strong> nhân viên từ file Excel gần nhất.</p>
      </div>
    ` : ""}
  `;

  res.send(renderLayout(html));
});

app.post("/upload", upload.single("excelFile"), (req, res) => {
  if (!req.file) {
    return res.status(400).send(renderLayout(`<div class="card"><h2>Lỗi</h2><p>Chưa có file Excel.</p><a class="btn" href="/">Quay lại</a></div>`));
  }

  const filePath = req.file.path;
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const headers = rows.length ? Object.keys(rows[0]) : [];

  const mapping = Object.keys(FIELD_LABELS).reduce((acc, field) => {
    acc[field] = guessColumn(headers, field);
    return acc;
  }, {});

  const employees = rows
    .map((row) => buildEmployee(row, mapping))
    .filter((item) => item.name);

  payrollData = { headers, rows, mapping, employees };

  fs.unlink(filePath, () => {});
  res.redirect("/employees");
});

app.get("/employees", (req, res) => {
  if (!payrollData.employees.length) {
    return res.redirect("/");
  }

  const search = normalize(req.query.q || "");
  const employees = payrollData.employees.filter((e) => !search || normalize(e.name).includes(search));

  const totalSalary = employees.reduce((sum, e) => sum + e.salary, 0);
  const totalNet = employees.reduce((sum, e) => sum + e.netIncome, 0);

  const html = `
    <div class="card no-print">
      <h1>Danh sách nhân viên</h1>
      <form method="GET" action="/employees">
        <input type="text" name="q" placeholder="Tìm theo tên nhân viên" value="${req.query.q || ""}" />
      </form>
      <br />
      <div class="grid grid-3">
        <div class="item"><div class="label">Số nhân viên</div><div class="value">${employees.length}</div></div>
        <div class="item"><div class="label">Tổng lương</div><div class="value">${money(totalSalary)}</div></div>
        <div class="item"><div class="label">Tổng thực lĩnh</div><div class="value">${money(totalNet)}</div></div>
      </div>
    </div>

    <div class="card">
      <table>
        <thead>
          <tr>
            <th>Tên</th>
            <th>Lương</th>
            <th>Thực lĩnh</th>
            <th class="no-print">Phiếu lương</th>
          </tr>
        </thead>
        <tbody>
          ${employees.map((e, i) => `
            <tr>
              <td>${e.name}</td>
              <td>${money(e.salary)}</td>
              <td>${money(e.netIncome)}</td>
              <td class="no-print"><a class="btn btn-outline" href="/slip/${i}">Xem phiếu</a></td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>

    <div class="no-print"><a class="btn" href="/">Tải file khác</a></div>
  `;

  res.send(renderLayout(html, "Danh sách nhân viên"));
});

app.get("/slip/:index", (req, res) => {
  const employee = payrollData.employees[Number(req.params.index)];
  if (!employee) {
    return res.status(404).send(renderLayout(`<div class="card"><h2>Không tìm thấy nhân viên</h2><a class="btn" href="/employees">Quay lại</a></div>`));
  }

  const html = `
    <div class="card">
      <div class="actions no-print">
        <a class="btn btn-outline" href="/employees">Quay lại</a>
        <button class="btn" onclick="window.print()">In phiếu</button>
      </div>
    </div>

    <div class="card">
      <h1>Phiếu lương nhân viên</h1>
      <p>Nhân viên: <strong>${employee.name}</strong></p>

      <div class="grid grid-2">
        <div class="item"><div class="label">Tên</div><div class="value">${employee.name}</div></div>
        <div class="item"><div class="label">Lương</div><div class="value">${money(employee.salary)}</div></div>
        <div class="item"><div class="label">Phụ cấp</div><div class="value">${money(employee.allowance)}</div></div>
        <div class="item"><div class="label">Thu nhập khác</div><div class="value">${money(employee.otherIncome)}</div></div>
        <div class="item"><div class="label">Thưởng/Sinh nhật</div><div class="value">${money(employee.birthdayBonus)}</div></div>
        <div class="item"><div class="label">Tăng ca</div><div class="value">${money(employee.overtime)}</div></div>
        <div class="item"><div class="label">Ngày công nghỉ</div><div class="value">${employee.daysOff || 0}</div></div>
        <div class="item"><div class="label">Trừ tạm ứng</div><div class="value">${money(employee.advanceDeduction)}</div></div>
        <div class="item"><div class="label">Trừ đã TT</div><div class="value">${money(employee.paidDeduction)}</div></div>
        <div class="item"><div class="label">BHXH</div><div class="value">${money(employee.socialInsurance)}</div></div>
        <div class="item"><div class="label">Thuế TNCN</div><div class="value">${money(employee.pitTax)}</div></div>
        <div class="item" style="background:#0f172a;color:white;"><div class="label" style="color:#cbd5e1;">Thực lĩnh</div><div class="value" style="font-size:24px;">${money(employee.netIncome)}</div></div>
      </div>
    </div>
  `;

  res.send(renderLayout(html, `Phiếu lương - ${employee.name}`));
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
