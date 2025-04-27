const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const port = 4000;
const PASSWORD = 'Kittiset3'; // รหัสผ่าน

const upload = multer({ dest: 'uploads/' });

app.use(express.static(__dirname));
app.use(bodyParser.urlencoded({ extended: true }));

const filePath = 'proposals.xlsx';
let workbook = new ExcelJS.Workbook();
let worksheet;

// ฟังก์ชันโหลดหรือสร้างไฟล์ Excel
async function loadWorkbook() {
  if (fs.existsSync(filePath)) {
    await workbook.xlsx.readFile(filePath);
    worksheet = workbook.getWorksheet(1);
    console.log("โหลดไฟล์ proposals.xlsx เรียบร้อยแล้ว");
  } else {
    worksheet = workbook.addWorksheet('ข้อเสนอแนะ');
    worksheet.columns = [
      { header: 'ลำดับ', key: 'no', width: 10 },
      { header: 'ชื่อ', key: 'name', width: 20 },
      { header: 'พื้นที่', key: 'area', width: 15 },
      { header: 'หัวข้อ', key: 'topic', width: 25 },
      { header: 'รายละเอียด', key: 'details', width: 50 },
    ];
    console.log("สร้างไฟล์ proposals.xlsx ใหม่");
  }
}

// ฟังก์ชันจัดรูปแบบเซลล์
function styleCell(cell, options = {}) {
  if (options.fill) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: options.fill },
    };
  }
  if (options.font) {
    cell.font = options.font;
  }
  if (options.alignment) {
    cell.alignment = options.alignment;
  }
  if (options.border) {
    cell.border = options.border;
  }
}

// รับข้อมูลฟอร์ม
app.post('/submit', upload.single('file'), async (req, res) => {
  const { name, area, topic, details } = req.body;

  await loadWorkbook();

  // ดึงข้อมูลเดิม
  let data = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      const rowData = {
        name: row.getCell(2).value,
        area: row.getCell(3).value,
        topic: row.getCell(4).value,
        details: row.getCell(5).value,
      };
      data.push(rowData);
    }
  });

  // เพิ่มข้อมูลใหม่
  data.push({ name, area, topic, details });

  // จัดกลุ่มตามพื้นที่
  const areas = ['โนนไทย', 'สูงเนิน', 'ขามทะเลสอ'];
  const groupedData = {};
  areas.forEach(a => groupedData[a] = []);
  data.forEach(item => {
    if (groupedData[item.area]) {
      groupedData[item.area].push(item);
    }
  });

  // ล้างข้อมูลเดิม
  worksheet.spliceRows(2, worksheet.rowCount);

  let rowIndex = 2;
  areas.forEach(area => {
    const items = groupedData[area];
    if (items.length > 0) {
      // แถวหัวพื้นที่
      worksheet.mergeCells(`A${rowIndex}:E${rowIndex}`);
      const headerCell = worksheet.getCell(`A${rowIndex}`);
      headerCell.value = area;
      styleCell(headerCell, {
        fill: 'FFD9D9D9', // สีเทาอ่อน
        font: { bold: true },
        alignment: { horizontal: 'center', vertical: 'middle' },
      });
      rowIndex++;

      // รายการข้อเสนอ
      items.forEach((item, index) => {
        const row = worksheet.getRow(rowIndex);
        row.getCell(1).value = index + 1;
        row.getCell(2).value = item.name;
        row.getCell(3).value = item.area;
        row.getCell(4).value = item.topic;
        row.getCell(5).value = item.details;
        row.getCell(5).alignment = { wrapText: true, vertical: 'middle' };
        rowIndex++;
      });

      // แสดงยอดรวม
      worksheet.mergeCells(`A${rowIndex}:E${rowIndex}`);
      const totalCell = worksheet.getCell(`A${rowIndex}`);
      totalCell.value = `รวม ${items.length} ข้อเสนอ`;
      styleCell(totalCell, {
        font: { italic: true },
        alignment: { horizontal: 'right', vertical: 'middle' },
      });
      rowIndex++;
    }
  });

  await workbook.xlsx.writeFile(filePath);

  res.send(`<h2>ส่งข้อเสนอเรียบร้อยแล้ว</h2><a href="/">กลับไปส่งอีกครั้ง</a>`);
});

// หน้า view ใส่รหัสผ่าน
app.get('/view', (req, res) => {
  res.send(`
    <h2>ใส่รหัสผ่านเพื่อดูข้อเสนอแนะ</h2>
    <form method="POST" action="/view">
      <input type="password" name="password" placeholder="รหัสผ่าน" required/>
      <button type="submit">เข้าสู่ระบบ</button>
    </form>
  `);
});

// ตรวจสอบรหัสผ่านแล้วแสดงข้อมูล
app.post('/view', async (req, res) => {
  const password = req.body.password;

  if (password === PASSWORD) {
    await loadWorkbook();
    worksheet = workbook.getWorksheet(1);

    let rows = worksheet.getSheetValues();
    rows = rows.slice(2);

    let table = `
      <h2>รายการข้อเสนอแนะ</h2>
      <table border="1" cellpadding="5" cellspacing="0">
        <tr>
          <th>ลำดับ</th>
          <th>ชื่อ</th>
          <th>พื้นที่</th>
          <th>หัวข้อ</th>
          <th>รายละเอียด</th>
        </tr>
    `;

    rows.forEach((row) => {
      if (row && row[1] && row[2] && row[3]) {
        table += `
          <tr>
            <td>${row[1]}</td>
            <td>${row[2]}</td>
            <td>${row[3]}</td>
            <td>${row[4]}</td>
            <td>${row[5]}</td>
          </tr>
        `;
      } else if (row && typeof row[1] === 'string' && !row[2]) {
        table += `
          <tr style="background-color: #f2f2f2;">
            <td colspan="5" align="center"><strong>${row[1]}</strong></td>
          </tr>
        `;
      }
    });

    table += `</table><br><a href="/">กลับหน้าแรก</a>`;

    res.send(table);
  } else {
    res.send(`<h2>รหัสผ่านไม่ถูกต้อง!</h2><a href="/view">ลองใหม่</a>`);
  }
});

// หน้าแรก
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'form.html'));
});

// เริ่มต้นเซิร์ฟเวอร์
app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});