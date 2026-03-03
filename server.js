const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const Database = require('better-sqlite3');

const app = express();
const db = new Database('expiry.db');

db.exec(`
  CREATE TABLE IF NOT EXISTS products (
    sample_no TEXT PRIMARY KEY,
    product_name TEXT,
    expiry_date TEXT,
    generate_date TEXT,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
  )
`);
db.exec(`CREATE INDEX IF NOT EXISTS idx_sample_no ON products(sample_no)`);

const IMPORT_PASSWORD = 'admin123';

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const upload = multer({ storage: multer.memoryStorage() });

app.post('/api/import', upload.single('file'), (req, res) => {
  if (req.body?.password !== IMPORT_PASSWORD) {
    return res.status(401).json({ error: '密码错误' });
  }
  if (!req.file) {
    return res.status(400).json({ error: '请上传文件' });
  }
  
  try {
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet);
    
    if (rows.length === 0) {
      return res.json({ success: true, count: 0 });
    }

    const keys = Object.keys(rows[0]);
    const sampleKey = keys.find(k => k.includes('样本编号') || k.toLowerCase().includes('sample')) || keys[0];
    const productKey = keys.find(k => k.includes('产品')) || keys[1];
    const expiryKey = keys.find(k => k.includes('效期') || k.toLowerCase().includes('expiry')) || keys[2];
    const generateKey = keys.find(k => k.includes('生产日期') || k.includes('生成日期')) || keys[3];

    const insert = db.prepare(`
      INSERT OR REPLACE INTO products (sample_no, product_name, expiry_date, generate_date, updated_at)
      VALUES (?, ?, ?, ?, datetime('now'))
    `);

    const insertMany = db.transaction((items) => {
      let count = 0;
      let emptyCount = 0;
      for (const item of items) {
        const sampleNo = String(item[sampleKey] || '').trim();
        const productName = String(item[productKey] || '').trim();
        const expiryDate = String(item[expiryKey] || '').trim();
        const generateDate = String(item[generateKey] || '').trim();
        if (sampleNo) {
          insert.run(sampleNo, productName, expiryDate, generateDate);
          count++;
        } else {
          emptyCount++;
        }
      }
      return { count, emptyCount };
    });

    const result = insertMany(rows);
    res.json({ success: true, count: result.count, empty: result.emptyCount });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/delete-by-file', upload.single('file'), (req, res) => {
  if (req.body?.password !== IMPORT_PASSWORD) {
    return res.status(401).json({ error: '密码错误' });
  }
  if (!req.file) {
    return res.status(400).json({ error: '请上传文件' });
  }
  
  try {
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet);
    
    if (rows.length === 0) {
      return res.json({ success: true, deleted: 0 });
    }

    const keys = Object.keys(rows[0]);
    const sampleKey = keys.find(k => k.includes('样本编号') || k.toLowerCase().includes('sample')) || keys[0];

    const sampleNos = rows.map(item => String(item[sampleKey] || '').trim()).filter(s => s);
    
    if (sampleNos.length === 0) {
      return res.status(400).json({ error: '文件中未找到样本编号' });
    }

    const placeholders = sampleNos.map(() => '?').join(',');
    const stmt = db.prepare(`DELETE FROM products WHERE sample_no IN (${placeholders})`);
    const result = stmt.run(...sampleNos);
    
    res.json({ success: true, deleted: result.changes });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/query/batch', (req, res) => {
  try {
    const { sampleNos } = req.body;
    if (!Array.isArray(sampleNos) || sampleNos.length === 0) {
      return res.json({ success: true, data: [] });
    }

    const placeholders = sampleNos.map(() => '?').join(',');
    const stmt = db.prepare(`SELECT sample_no, product_name, expiry_date, generate_date FROM products WHERE sample_no IN (${placeholders})`);
    const rows = stmt.all(...sampleNos);

    const result = sampleNos.map(no => {
      const found = rows.find(r => r.sample_no === no);
      return found || { sample_no: no, product_name: '', expiry_date: '未找到', generate_date: '' };
    });

    res.json({ success: true, data: result });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/export', (req, res) => {
  try {
    const { data } = req.body;
    if (!Array.isArray(data) || data.length === 0) {
      return res.status(400).json({ error: '无数据可导出' });
    }

    const worksheet = xlsx.utils.json_to_sheet(data);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, '效期查询结果');
    const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', "attachment; filename*=UTF-8''" + encodeURIComponent('效期查询结果.xlsx'));
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/delete', (req, res) => {
  try {
    if (req.body?.password !== IMPORT_PASSWORD) {
      return res.status(401).json({ error: '密码错误' });
    }
    const { sampleNos } = req.body;
    
    if (!sampleNos || !Array.isArray(sampleNos) || sampleNos.length === 0) {
      return res.status(400).json({ error: '请提供要删除的样本编号' });
    }
    
    const placeholders = sampleNos.map(() => '?').join(',');
    const stmt = db.prepare(`DELETE FROM products WHERE sample_no IN (${placeholders})`);
    const result = stmt.run(...sampleNos);
    res.json({ success: true, deleted: result.changes });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/stats', (req, res) => {
  const count = db.prepare('SELECT COUNT(*) as count FROM products').get();
  res.json(count);
});

const PORT = 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
