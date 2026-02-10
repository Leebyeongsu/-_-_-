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
let selectedExcelFile = null;

// 엑셀 변환 관련 요소
const excelDropZone = document.getElementById('excel-drop-zone');
const excelInput = document.getElementById('excel-input');
const excelPreview = document.getElementById('excel-preview');
const excelFilename = document.getElementById('excel-filename');
const excelFilesize = document.getElementById('excel-filesize');
const excelConvertBtn = document.getElementById('excel-convert-btn');
const excelConvertText = document.getElementById('excel-convert-text');

// 색상 팔레트 (원본 현황표와 동일)
const COLOR_MAP = {
  'YELLOW': { rgb: "FFFF99", html: '#FFFFCC' },
  'GREEN': { rgb: "C6EFCE", html: '#C6EFCE' },
  'PINK': { rgb: "FFCCFF", html: '#FFCCFF' },
  'WHITE': { rgb: "FFFFFF", html: '#FFFFFF' }
};

function init() {
  // 이미지 변환
  dropZone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));
  convertBtn.addEventListener('click', startConversion);
  downloadBtn.addEventListener('click', downloadExcel);

  // 엑셀 변환
  excelDropZone.addEventListener('click', () => excelInput.click());
  excelInput.addEventListener('change', (e) => handleExcelFile(e.target.files[0]));
  excelConvertBtn.addEventListener('click', startExcelConversion);
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
    const response = await fetch('/api/ocr', {
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

async function downloadExcel() {
  if (!processedData || processedData.length === 0) return;

  downloadBtn.disabled = true;
  const originalText = downloadBtn.textContent;
  downloadBtn.textContent = '생성 중...';

  try {
    // 서버에 JSON 데이터 전송하여 엑셀 생성
    const response = await fetch('/api/download-basic-excel', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ data: processedData })
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error || '서버 오류');
    }

    // 서버에서 생성한 엑셀 파일 다운로드
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;

    // 파일명 생성
    const headerEl = document.getElementById('header-name');
    const buildingEl = document.getElementById('header-building');
    const aptName = headerEl?.textContent || '현황표';
    const dongName = buildingEl?.textContent || '';
    const fileName = dongName ? `${aptName}_${dongName}` : aptName;
    a.download = `${fileName}_${new Date().getTime()}.xlsx`;

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);

    downloadBtn.textContent = '완료! ✅';
    setTimeout(() => {
      downloadBtn.textContent = originalText;
      downloadBtn.disabled = false;
    }, 2000);

  } catch (error) {
    alert(`다운로드 실패: ${error.message}`);
    downloadBtn.textContent = originalText;
    downloadBtn.disabled = false;
  }
}

// 엑셀 파일 처리
function handleExcelFile(file) {
  if (!file) return;
  selectedExcelFile = file;

  // 파일 정보 표시
  excelFilename.textContent = file.name;
  const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
  excelFilesize.textContent = `${sizeInMB} MB`;

  excelPreview.classList.remove('hidden');
  excelConvertBtn.disabled = false;
}

// 엑셀 변환 시작
async function startExcelConversion() {
  if (!selectedExcelFile) return;

  excelConvertBtn.disabled = true;
  excelConvertText.innerHTML = '<span class="loader"></span>변환 중...';

  const formData = new FormData();
  formData.append('excel', selectedExcelFile);

  try {
    console.log('📤 엑셀 파일 업로드 중:', selectedExcelFile.name);

    const response = await fetch('/api/convert-excel', {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error || '서버 오류');
    }

    console.log('✅ 변환 완료, 다운로드 시작');

    // 변환된 엑셀 파일 다운로드
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;

    // 파일명 생성 (원본명_변환.xlsx)
    const originalName = selectedExcelFile.name.replace(/\.(xlsx|xls)$/i, '');
    a.download = `${originalName}_변환.xlsx`;

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);

    console.log('✅ 다운로드 완료');

    excelConvertText.textContent = '변환 완료! ✅';
    setTimeout(() => {
      excelConvertText.textContent = '층호수 형태로 변환';
      excelConvertBtn.disabled = false;
    }, 2000);

  } catch (error) {
    console.error('❌ 변환 실패:', error);
    alert(`변환 실패: ${error.message}\n\n명령줄 방식을 사용하세요:\npython excel_converter.py "파일명.xlsx"`);
    excelConvertText.textContent = '층호수 형태로 변환';
    excelConvertBtn.disabled = false;
  }
}

init();
