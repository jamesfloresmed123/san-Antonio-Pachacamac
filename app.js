const tabButtons = document.querySelectorAll('.tab-button');
const tabPanels = document.querySelectorAll('.tab-panel');

const qrForm = document.getElementById('qr-form');
const qrContainer = document.getElementById('qrcode');
const jsonPreview = document.getElementById('json-preview');

const excelStatus = document.getElementById('excel-status');
const decodedDataEl = document.getElementById('decoded-data');
const validationResultEl = document.getElementById('validation-result');
const startScanBtn = document.getElementById('start-scan');
const stopScanBtn = document.getElementById('stop-scan');

const DEFAULT_ASSOCIATES_FILE = 'data/asociados.csv';
const ASSOCIATE_INVALID_MESSAGE = 'asociado inexistente o no está la día';
const DEFAULT_ASSOCIATES_FALLBACK = [
  {
    'Código asociado': '130',
    'Nombre completo': 'James Abraham Flores Medina',
    Etapa: '8',
    Estado: 'Activo',
  },
];

let asociados = [];
let scanner = null;
let scannerRunning = false;

function normalizeText(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function setValidationMessage(message, type = '') {
  validationResultEl.textContent = message;
  validationResultEl.classList.remove('success', 'error');
  if (type) validationResultEl.classList.add(type);
}

function parseExcelRow(row) {
  const normalized = {};
  Object.keys(row).forEach((key) => {
    normalized[normalizeText(key)] = row[key];
  });

  const codigo =
    normalized['codigo'] ??
    normalized['codigo asociado'] ??
    normalized['cod asociado'] ??
    normalized['id asociado'];

  const estado = normalized['estado'] ?? normalized['estatus'] ?? normalized['status'];

  return {
    codigo: String(codigo ?? '').trim(),
    estado: String(estado ?? '').trim(),
    raw: row,
  };
}

function isActiveStatus(status) {
  const value = normalizeText(status);
  const activeStates = new Set(['activo', 'activa', 'habilitado', 'habilitada', 'si', '1', 'true']);
  return activeStates.has(value);
}

function validateAssociate(decodedObj) {
  if (!asociados.length) {
    setValidationMessage('No se pudo cargar la base de asociados.', 'error');
    return;
  }

  const codigoBuscado = String(decodedObj.codigoAsociado ?? '').trim();
  if (!codigoBuscado) {
    setValidationMessage('El QR no contiene el código de asociado.', 'error');
    return;
  }

  const asociado = asociados.find((item) => item.codigo === codigoBuscado);
  if (!asociado || !isActiveStatus(asociado.estado)) {
    setValidationMessage(ASSOCIATE_INVALID_MESSAGE, 'error');
    return;
  }

  setValidationMessage(`Asociado ${codigoBuscado} válido y ACTIVO.`, 'success');
}

function decodeQrText(text) {
  let parsed = null;
  try {
    parsed = JSON.parse(text);
  } catch {
    setValidationMessage('El QR leído no tiene formato JSON válido.', 'error');
    decodedDataEl.textContent = text;
    return;
  }

  decodedDataEl.textContent = JSON.stringify(parsed, null, 2);
  validateAssociate(parsed);
}

function parseCsvRows(content) {
  const lines = String(content)
    .replace(/^\uFEFF/, '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) return [];

  const headers = lines[0].split(',').map((header) => header.trim());

  return lines.slice(1).map((line) => {
    const values = line.split(',').map((value) => value.trim());
    return headers.reduce((acc, header, index) => {
      acc[header] = values[index] ?? '';
      return acc;
    }, {});
  });
}

async function fetchWithTimeout(url, timeoutMs = 7000) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

  try {
    return await fetch(url, { signal: controller.signal });
  } finally {
    clearTimeout(timeoutId);
  }
}

async function loadAssociatesFile() {
  try {
    const response = await fetchWithTimeout(DEFAULT_ASSOCIATES_FILE);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);

    const fileContent = await response.text();
    const rows = parseCsvRows(fileContent);

    asociados = rows.map(parseExcelRow).filter((item) => item.codigo);

    if (!asociados.length) {
      throw new Error('archivo sin registros válidos');
    }

    excelStatus.textContent = `Base cargada desde ${DEFAULT_ASSOCIATES_FILE}. Registros con código: ${asociados.length}.`;
  } catch (error) {
    asociados = DEFAULT_ASSOCIATES_FALLBACK.map(parseExcelRow).filter((item) => item.codigo);

    excelStatus.textContent = `No se pudo leer ${DEFAULT_ASSOCIATES_FILE}; usando base incluida. Registros con código: ${asociados.length}.`;
    setValidationMessage('Usando base de respaldo local para validar asociados.');
  }
}

function stopScanner() {
  if (scanner && scannerRunning) {
    scanner.stop().then(() => {
      scannerRunning = false;
      startScanBtn.disabled = false;
      stopScanBtn.disabled = true;
    });
  }
}

tabButtons.forEach((button) => {
  button.addEventListener('click', () => {
    tabButtons.forEach((btn) => btn.classList.remove('active'));
    tabPanels.forEach((panel) => panel.classList.remove('active'));

    button.classList.add('active');
    document.getElementById(button.dataset.tab).classList.add('active');
  });
});

qrForm.addEventListener('submit', (event) => {
  event.preventDefault();

  const payload = {
    codigoAsociado: document.getElementById('codigo').value.trim(),
    nombreCompleto: document.getElementById('nombre').value.trim(),
    etapa: document.getElementById('etapa').value.trim(),
    generadoEn: new Date().toISOString(),
  };

  qrContainer.innerHTML = '';
  new QRCode(qrContainer, {
    text: JSON.stringify(payload),
    width: 220,
    height: 220,
  });

  jsonPreview.textContent = JSON.stringify(payload, null, 2);
});

startScanBtn.addEventListener('click', async () => {
  try {
    if (!scanner) scanner = new Html5Qrcode('reader');

    await scanner.start(
      { facingMode: 'environment' },
      { fps: 10, qrbox: { width: 240, height: 240 } },
      (decodedText) => {
        decodeQrText(decodedText);
        stopScanner();
      },
      () => {},
    );

    scannerRunning = true;
    startScanBtn.disabled = true;
    stopScanBtn.disabled = false;
    setValidationMessage('Escaneando... apunta al código QR.');
  } catch (error) {
    setValidationMessage(`No se pudo iniciar la cámara: ${error.message}`, 'error');
  }
});

stopScanBtn.addEventListener('click', () => {
  stopScanner();
  setValidationMessage('Escaneo detenido.');
});

loadAssociatesFile();
