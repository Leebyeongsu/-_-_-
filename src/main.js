import './style.css'
import XLSX from 'xlsx-js-style';

const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const imagePreview = document.getElementById('image-preview');
const convertBtn = document.getElementById('convert-btn');
const convertText = document.getElementById('convert-text');
const downloadBtn = document.getElementById('download-btn');
const statusBadge = document.getElementById('status-badge');
const resultsPlaceholder = document.getElementById('results-placeholder');
const tableContainer = document.getElementById('table-container');
const dataTable = document.getElementById('data-table');
const progressContainer = document.getElementById('progress-container');
const progressBar = document.getElementById('progress-bar');

let processedData = null;
let selectedFile = null;

// 색상 팔레트 (원본 현황표와 동일)
const COLOR_MAP = {
  'YELLOW': { rgb: "FFFF99", html: '#FFFFCC' },
  'GREEN': { rgb: "C6EFCE", html: '#C6EFCE' },
  'PINK': { rgb: "FFCCFF", html: '#FFCCFF' },
  'WHITE': { rgb: "FFFFFF", html: '#FFFFFF' }
};

function init() {
  dropZone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));
  convertBtn.addEventListener('click', startConversion);
  downloadBtn.addEventListener('click', downloadExcel);
}

function handleFile(file) {
  if (!file) return;
  selectedFile = file;
  const reader = new FileReader();
  reader.onload = (e) => {
    imagePreview.src = e.target.result;
    imagePreview.style.display = 'block';
    resultsPlaceholder.style.display = 'none';
    convertBtn.disabled = false;
    updateStatus('이미지 로드 완료', 'idle');
  };
  reader.readAsDataURL(file);
}

function updateStatus(text, type) {
  statusBadge.textContent = text;
  statusBadge.className = `status-badge status-${type}`;
}

async function startConversion() {
  if (!selectedFile) return;

  convertBtn.disabled = true;
  downloadBtn.disabled = true;
  convertText.innerHTML = '<span class="loader"></span>AI 분석 중...';
  updateStatus('분석 중', 'processing');
  progressContainer.classList.remove('hidden');
  progressBar.style.width = '30%';

  const formData = new FormData();
  formData.append('image', selectedFile);

  try {
    const response = await fetch('http://localhost:3500/api/ocr', {
      method: 'POST',
      body: formData
    });

    if (!response.ok) throw new Error('서버 오류');

    const rawData = await response.json();
    // 새 포맷: { header, data } 또는 기존 배열 호환
    const floorData = rawData.data || rawData;
    const headerInfo = rawData.header || null;
    processedData = floorData;

    progressBar.style.width = '70%';
    renderHeader(headerInfo);
    renderTable(floorData);

    progressBar.style.width = '100%';
    updateStatus('분석 완료', 'done');
    convertText.textContent = '다시 변환';
    convertBtn.disabled = false;
    downloadBtn.disabled = false;

    setTimeout(() => progressContainer.classList.add('hidden'), 500);
  } catch (error) {
    updateStatus(`실패: ${error.message}`, 'idle');
    convertText.textContent = '변환 시작';
    convertBtn.disabled = false;
    progressContainer.classList.add('hidden');
  }
}

function renderHeader(headerInfo) {
  const headerEl = document.getElementById('header-info');
  const buildingEl = document.getElementById('header-building');
  const nameEl = document.getElementById('header-name');

  if (!headerInfo || (!headerInfo.building && !headerInfo.name)) {
    headerEl.classList.add('hidden');
    return;
  }

  headerEl.classList.remove('hidden');
  buildingEl.textContent = headerInfo.building || '';
  buildingEl.style.display = headerInfo.building ? 'inline-block' : 'none';
  nameEl.textContent = headerInfo.name || '';
  nameEl.style.display = headerInfo.name ? 'inline-block' : 'none';
}

function renderTable(data) {
  dataTable.innerHTML = '';
  tableContainer.classList.remove('hidden');
  resultsPlaceholder.classList.add('hidden');

  if (!data || data.length === 0) {
    alert('데이터가 없습니다.');
    return;
  }

  // 첫 번째 행에서 호수 목록 추출
  const firstRow = data[0];
  const unitKeys = Object.keys(firstRow.units || {}).sort((a, b) => {
    const numA = parseInt(a.replace(/[^0-9]/g, '')) || 0;
    const numB = parseInt(b.replace(/[^0-9]/g, '')) || 0;
    return numA - numB;
  });

  // 헤더 생성
  const trHead = document.createElement('tr');
  let headerHtml = '<th>층</th>';
  unitKeys.forEach(key => {
    headerHtml += `<th>${key}</th>`;
  });
  trHead.innerHTML = headerHtml;
  dataTable.appendChild(trHead);

  // 데이터 행 생성
  data.forEach(item => {
    const tr = document.createElement('tr');

    // 층수
    const tdFloor = document.createElement('td');
    tdFloor.textContent = item.floor || '';
    tr.appendChild(tdFloor);

    // 각 호수
    unitKeys.forEach(unitKey => {
      const td = document.createElement('td');
      const cellData = item.units?.[unitKey] || { text: '', color: 'WHITE' };

      // 텍스트만 표시 (색상은 배경으로만)
      td.textContent = cellData.text || '';

      // 배경색 적용
      const colorKey = String(cellData.color || 'WHITE').toUpperCase();
      if (COLOR_MAP[colorKey]) {
        td.style.backgroundColor = COLOR_MAP[colorKey].html;
      }

      tr.appendChild(td);
    });

    dataTable.appendChild(tr);
  });
}

function downloadExcel() {
  if (!processedData || processedData.length === 0) return;

  const wb = XLSX.utils.book_new();

  const baseStyle = {
    border: {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' }
    },
    alignment: { horizontal: 'center', vertical: 'center' }
  };

  // 첫 번째 행에서 호수 목록 추출
  const firstRow = processedData[0];
  const unitKeys = Object.keys(firstRow.units || {}).sort((a, b) => {
    const numA = parseInt(a.replace(/[^0-9]/g, '')) || 0;
    const numB = parseInt(b.replace(/[^0-9]/g, '')) || 0;
    return numA - numB;
  });

  const wsData = [];

  // 헤더
  const headerRow = [
    { v: '층', s: { ...baseStyle, font: { bold: true }, fill: { fgColor: { rgb: "DDDDDD" } } } }
  ];
  unitKeys.forEach(key => {
    headerRow.push({ v: key, s: { ...baseStyle, font: { bold: true }, fill: { fgColor: { rgb: "DDDDDD" } } } });
  });
  wsData.push(headerRow);

  // 데이터 행
  processedData.forEach(item => {
    const row = [
      { v: item.floor || '', s: baseStyle }
    ];

    unitKeys.forEach(unitKey => {
      const cellData = item.units?.[unitKey] || { text: '', color: 'WHITE' };
      const colorKey = String(cellData.color || 'WHITE').toUpperCase();
      const colorHex = COLOR_MAP[colorKey]?.rgb || "FFFFFF";

      row.push({
        v: cellData.text || '',
        s: {
          ...baseStyle,
          fill: { fgColor: { rgb: colorHex } }
        }
      });
    });

    wsData.push(row);
  });

  // 워크시트 생성
  const ws = XLSX.utils.aoa_to_sheet([]);
  wsData.forEach((row, rIdx) => {
    row.forEach((cell, cIdx) => {
      const addr = XLSX.utils.encode_cell({ r: rIdx, c: cIdx });
      ws[addr] = cell;
    });
  });

  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: wsData.length - 1, c: unitKeys.length }
  });

  XLSX.utils.book_append_sheet(wb, ws, "현황분석");

  // 파일명: 헤더 정보가 있으면 활용, 없으면 기본값
  const headerEl = document.getElementById('header-name');
  const buildingEl = document.getElementById('header-building');
  const aptName = headerEl?.textContent || '현황표';
  const dongName = buildingEl?.textContent || '';
  const fileName = dongName ? `${aptName}_${dongName}` : aptName;
  XLSX.writeFile(wb, `${fileName}_${new Date().getTime()}.xlsx`);
}

init();
