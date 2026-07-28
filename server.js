const http = require('http');
const fs = require('fs');
const path = require('path');
const XLSX = require('./xlsx.full.min.js');

const root = __dirname;
const port = Number(process.env.PORT || 5173);
const host = process.env.HOST || '0.0.0.0';
const statePath = path.join(root, 'cuotas_state.json');
const types = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
};

function sendJson(res, status, payload) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(payload));
}

function readJsonBody(req, onDone, onError) {
  let body = '';
  req.on('data', (chunk) => {
    body += chunk;
    if (body.length > 10_000_000) req.destroy();
  });
  req.on('end', () => {
    try {
      onDone(JSON.parse(body || '{}'));
    } catch (err) {
      onError(err);
    }
  });
}

function formatearMonedaExcel(ws, rows, header) {
  const range = XLSX.utils.decode_range(ws['!ref']);
  let col = -1;
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
    if (cell && cell.v === header) col = c;
  }
  if (col < 0) return;
  for (let r = 1; r <= rows.length; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: col });
    if (ws[addr]) {
      ws[addr].t = 'n';
      ws[addr].z = '"$" #,##0.00';
    }
  }
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);

  if (req.method === 'GET' && url.pathname === '/api/state') {
    try {
      if (!fs.existsSync(statePath)) {
        sendJson(res, 200, {});
        return;
      }
      const raw = fs.readFileSync(statePath, 'utf8');
      sendJson(res, 200, JSON.parse(raw || '{}'));
    } catch (err) {
      sendJson(res, 500, { error: err.message });
    }
    return;
  }

  if (req.method === 'POST' && url.pathname === '/api/state') {
    readJsonBody(req, (payload) => {
      try {
        const data = {
          smvm: Number(payload.smvm || 0),
          periodo: String(payload.periodo || ''),
          base: Array.isArray(payload.base) ? payload.base : [],
          reportes: payload.reportes && typeof payload.reportes === 'object' ? payload.reportes : {},
          updatedAt: new Date().toISOString()
        };
        fs.writeFileSync(statePath, JSON.stringify(data, null, 2), 'utf8');
        sendJson(res, 200, { ok: true, updatedAt: data.updatedAt });
      } catch (err) {
        sendJson(res, 500, { error: err.message });
      }
    }, (err) => sendJson(res, 400, { error: err.message }));
    return;
  }

  if (req.method === 'POST' && url.pathname === '/export-general') {
    readJsonBody(req, (payload) => {
      try {
        const rows = Array.isArray(payload.rows) ? payload.rows : [];
        if (!rows.length) {
          sendJson(res, 400, { error: 'Sin filas para exportar' });
          return;
        }

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(rows);
        ws['!cols'] = [
          { wch: 18 }, { wch: 32 }, { wch: 12 }, { wch: 16 }, { wch: 28 },
          { wch: 14 }, { wch: 16 }, { wch: 18 }, { wch: 42 },
        ];
        formatearMonedaExcel(ws, rows, 'Monto a depositar');
        XLSX.utils.book_append_sheet(wb, ws, 'Reporte general');

        const exportsDir = path.join(root, 'exports');
        fs.mkdirSync(exportsDir, { recursive: true });
        const safePeriodo = String(payload.periodo || 'periodo')
          .replace(/[\\/:*?"<>|]+/g, '-')
          .replace(/\s+/g, '_');
        const stamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 12);
        const filePath = path.join(exportsDir, `reporte_general_cuotas_${safePeriodo}_${stamp}.xlsx`);
        const bytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        fs.writeFileSync(filePath, Buffer.from(bytes));

        sendJson(res, 200, { path: filePath });
      } catch (err) {
        sendJson(res, 500, { error: err.message });
      }
    }, (err) => sendJson(res, 400, { error: err.message }));
    return;
  }

  const requested = url.pathname === '/' ? 'index.html' : decodeURIComponent(url.pathname.slice(1));
  const filePath = path.resolve(root, requested);

  if (!filePath.startsWith(root)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found');
      return;
    }
    res.writeHead(200, { 'Content-Type': types[path.extname(filePath)] || 'application/octet-stream' });
    res.end(data);
  });
});

server.listen(port, host, () => {
  console.log(`Cuotas alimentarias local: http://127.0.0.1:${port}`);
  const nets = require('os').networkInterfaces();
  Object.values(nets).flat().filter(x => x && x.family === 'IPv4' && !x.internal).forEach(x => {
    console.log(`Cuotas alimentarias red:   http://${x.address}:${port}`);
  });
});
