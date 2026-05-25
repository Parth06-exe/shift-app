const { app } = require('@azure/functions');
const { BlobServiceClient } = require('@azure/storage-blob');
const XLSX = require('xlsx');

const DEFAULT_CONTAINER = 'schedules';
const DEFAULT_BLOB = 'shift.xlsx';

function responseHeaders() {
  return {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': process.env.ALLOWED_ORIGIN || '*',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
}

function json(status, body) {
  return { status, headers: responseHeaders(), jsonBody: body };
}

function monthFromTitle(title) {
  const months = {
    january: 1, jan: 1, januar: 1,
    february: 2, feb: 2, februar: 2,
    march: 3, mar: 3, marec: 3,
    april: 4, apr: 4,
    may: 5, maj: 5,
    june: 6, jun: 6, junij: 6,
    july: 7, jul: 7, julij: 7,
    august: 8, aug: 8, avgust: 8,
    september: 9, sep: 9,
    october: 10, oct: 10, oktober: 10,
    november: 11, nov: 11,
    december: 12, dec: 12
  };
  const lower = String(title || '').toLowerCase();
  for (const [key, value] of Object.entries(months)) {
    if (lower.includes(key)) return value;
  }
  return new Date().getMonth() + 1;
}

function isWeekdayRow(row, nameCol) {
  const words = new Set(['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun', 'pon', 'tor', 'sre', 'cet', 'čet', 'pet', 'sob', 'ned']);
  let hits = 0;
  for (let i = nameCol + 1; i < row.length; i++) {
    if (words.has(String(row[i] || '').trim().toLowerCase())) hits++;
  }
  return hits >= 3;
}

function parseWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!worksheet || !worksheet['!ref']) return null;

  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = worksheet[XLSX.utils.encode_cell({ r, c })];
      row.push(cell ? String(cell.v).trim() : '');
    }
    rows.push(row);
  }
  if (!rows.length) return null;

  const title = rows[0].join(' ');
  const monthNum = monthFromTitle(title);
  const yearMatch = title.match(/\d{4}/);
  const year = yearMatch ? parseInt(yearMatch[0], 10) : new Date().getFullYear();

  let headerRow = -1;
  let nameCol = -1;
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      if (String(rows[r][c]).toLowerCase() === 'ime') {
        headerRow = r;
        nameCol = c;
        break;
      }
    }
    if (headerRow >= 0) break;
  }
  if (headerRow < 0) return null;

  const dayRow = rows[headerRow];
  let codeCol = -1;
  for (let c = 0; c < dayRow.length; c++) {
    const label = String(dayRow[c] || '').toLowerCase();
    if (['code', 'pin', 'employee code', 'sifra', 'šifra'].includes(label)) {
      codeCol = c;
      break;
    }
  }

  let firstPersonRow = headerRow + 1;
  if (rows[firstPersonRow] && isWeekdayRow(rows[firstPersonRow], nameCol)) firstPersonRow++;

  const dayColumns = [];
  for (let c = nameCol + 1; c < dayRow.length; c++) {
    const day = parseInt(dayRow[c], 10);
    if (day) dayColumns.push({ day, col: c });
  }

  const people = [];
  for (let r = firstPersonRow; r < rows.length; r++) {
    const row = rows[r];
    const name = String(row[nameCol] || '').trim();
    if (!name || name.startsWith('---')) continue;

    const shifts = {};
    for (const { day, col } of dayColumns) {
      const value = String(row[col] || '').trim();
      const parts = value.split(/\s+/);
      if (parts.length >= 2 && parts[0].toUpperCase() !== 'X' && parts[1].toUpperCase() !== 'X') {
        const start = parseInt(parts[0], 10);
        const end = parseInt(parts[1], 10);
        shifts[day] = !Number.isNaN(start) && !Number.isNaN(end)
          ? { start, end, overnight: start >= 13 && end <= 6, off: false }
          : { off: true };
      } else {
        shifts[day] = { off: true };
      }
    }

    people.push({
      name,
      code: codeCol >= 0 ? String(row[codeCol] || '').trim() : '',
      shifts
    });
  }

  return { people, monthNum, year };
}

function matchesName(personName, query) {
  const name = personName.toLowerCase();
  const search = query.toLowerCase();
  if (name.includes(search)) return true;
  return search.split(' ').some((word) => word.length > 1 && name.includes(word));
}

function toShiftList(person, monthNum, year) {
  return Object.keys(person.shifts)
    .map(Number)
    .sort((a, b) => a - b)
    .map((day) => ({ day, monthNum, year, ...person.shifts[day] }));
}

async function readBody(request) {
  try {
    return await request.json();
  } catch {
    return {};
  }
}

app.http('lookup', {
  methods: ['GET', 'POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: responseHeaders() };

    const body = request.method === 'POST' ? await readBody(request) : {};
    const name = String(request.query.get('name') || body.name || '').trim();
    const code = String(request.query.get('code') || body.code || '').trim();
    const requireCode = String(process.env.REQUIRE_EMPLOYEE_CODE || '').toLowerCase() === 'true';

    if (!name) return json(400, { error: 'Name is required.' });
    if (requireCode && !code) return json(400, { error: 'Employee code is required.' });

    const connectionString = process.env.SHIFT_STORAGE_CONNECTION_STRING || process.env.AzureWebJobsStorage;
    if (!connectionString) return json(500, { error: 'Storage connection string is not configured.' });

    try {
      const container = process.env.SHIFT_CONTAINER || DEFAULT_CONTAINER;
      const blob = process.env.SHIFT_BLOB || DEFAULT_BLOB;
      const blobClient = BlobServiceClient
        .fromConnectionString(connectionString)
        .getContainerClient(container)
        .getBlockBlobClient(blob);
      const buffer = await blobClient.downloadToBuffer();
      const schedule = parseWorkbook(buffer);
      if (!schedule || !schedule.people.length) return json(500, { error: 'Schedule file format is invalid.' });

      const person = schedule.people.find((candidate) => {
        if (!matchesName(candidate.name, name)) return false;
        if (!requireCode) return true;
        return candidate.code && candidate.code === code;
      });

      if (!person) return json(404, { error: 'Name not found in the schedule.' });

      return json(200, {
        name: person.name,
        monthNum: schedule.monthNum,
        year: schedule.year,
        shifts: toShiftList(person, schedule.monthNum, schedule.year)
      });
    } catch (error) {
      context.error(error);
      return json(500, { error: 'Could not read the schedule file.' });
    }
  }
});
