import './style.css'
import XLSX from 'xlsx-js-style';
import guideMarkdown from '../웹_사용_가이드.md?raw';

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

// 맵자료 수정 관련 요소
const mapDropZone = document.getElementById('map-drop-zone');
const mapInput = document.getElementById('map-input');
const mapPreview = document.getElementById('map-preview');
const mapFilename = document.getElementById('map-filename');
const mapFilesize = document.getElementById('map-filesize');
const mapConvertBtn = document.getElementById('map-convert-btn');
const mapConvertText = document.getElementById('map-convert-text');

// 서부산지사 맵자료 수정 관련 요소
const seoDropZone = document.getElementById('seo-drop-zone');
const seoInput = document.getElementById('seo-input');
const seoPreview = document.getElementById('seo-preview');
const seoFilename = document.getElementById('seo-filename');
const seoFilesize = document.getElementById('seo-filesize');
const seoConvertBtn = document.getElementById('seo-convert-btn');
const seoConvertText = document.getElementById('seo-convert-text');

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

  // 맵자료 수정
  mapDropZone.addEventListener('click', () => mapInput.click());
  mapInput.addEventListener('change', (e) => handleMapFile(e.target.files[0]));
  mapDropZone.addEventListener('dragover', (e) => { e.preventDefault(); mapDropZone.classList.add('active'); });
  mapDropZone.addEventListener('dragleave', () => mapDropZone.classList.remove('active'));
  mapDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    mapDropZone.classList.remove('active');
    const file = e.dataTransfer.files[0];
    if (file) handleMapFile(file);
  });
  mapConvertBtn.addEventListener('click', startMapConversion);

  // 서부산지사 맵자료 수정
  seoDropZone.addEventListener('click', () => seoInput.click());
  seoInput.addEventListener('change', (e) => handleSeoFile(e.target.files[0]));
  seoDropZone.addEventListener('dragover', (e) => { e.preventDefault(); seoDropZone.classList.add('active'); });
  seoDropZone.addEventListener('dragleave', () => seoDropZone.classList.remove('active'));
  seoDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    seoDropZone.classList.remove('active');
    const file = e.dataTransfer.files[0];
    if (file) handleSeoFile(file);
  });
  seoConvertBtn.addEventListener('click', startSeoConversion);

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
  // 한글 파일명이 Linux 서버에서 멀티파트 파싱 에러를 일으키므로 안전한 파일명으로 전송
  const imageExt = selectedFile.name.split('.').pop() || 'jpg';
  formData.append('image', selectedFile, `upload.${imageExt}`);

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

// 맵자료 파일 처리
function handleMapFile(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls'].includes(ext)) {
    alert('엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.');
    return;
  }

  mapFilename.textContent = file.name;
  const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
  mapFilesize.textContent = `${sizeInMB} MB`;
  mapPreview.classList.remove('hidden');
  mapConvertBtn.disabled = false;
  mapConvertBtn._file = file;
}

// 맵자료 수정 및 변환 시작
async function startMapConversion() {
  const file = mapConvertBtn._file;
  if (!file) return;

  mapConvertBtn.disabled = true;
  mapConvertText.innerHTML = '<span class="loader"></span>처리 중...';

  const formData = new FormData();
  // 한글 파일명이 Linux 서버에서 멀티파트 파싱 에러를 일으키므로 안전한 파일명으로 전송
  formData.append('mapExcel', file, 'upload.xlsx');

  try {
    const response = await fetch('/api/convert-map', {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      let errorMessage = `서버 오류 (${response.status})`;
      try {
        const errorData = await response.json();
        errorMessage = errorData.error || errorMessage;
      } catch (_) { /* 빈 응답이거나 JSON이 아닌 경우 */ }
      throw new Error(errorMessage);
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const originalName = file.name.replace(/\.(xlsx|xls)$/i, '');
    a.download = `${originalName}_동별시트.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);

    mapConvertText.textContent = '변환 완료! ✅';
    setTimeout(() => {
      mapConvertText.textContent = '수정 및 변환';
      mapConvertBtn.disabled = false;
    }, 2000);

  } catch (error) {
    alert(`변환 실패: ${error.message}`);
    mapConvertText.textContent = '수정 및 변환';
    mapConvertBtn.disabled = false;
  }
}

// 서부산지사 파일 처리
function handleSeoFile(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls'].includes(ext)) {
    alert('엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.');
    return;
  }
  seoFilename.textContent = file.name;
  const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
  seoFilesize.textContent = `${sizeInMB} MB`;
  seoPreview.classList.remove('hidden');
  seoConvertBtn.disabled = false;
  seoConvertBtn._file = file;
}

// 서부산지사 수정 및 변환 시작
async function startSeoConversion() {
  const file = seoConvertBtn._file;
  if (!file) return;

  seoConvertBtn.disabled = true;
  seoConvertText.innerHTML = '<span class="loader"></span>처리 중...';

  const formData = new FormData();
  formData.append('mapExcel', file, 'upload.xlsx');

  try {
    const response = await fetch('/api/convert-map-seo', {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      let errorMessage = `서버 오류 (${response.status})`;
      try {
        const errorData = await response.json();
        errorMessage = errorData.error || errorMessage;
      } catch (_) { }
      throw new Error(errorMessage);
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const originalName = file.name.replace(/\.(xlsx|xls)$/i, '');
    a.download = `${originalName}_동별시트.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);

    seoConvertText.textContent = '변환 완료! ✅';
    setTimeout(() => {
      seoConvertText.textContent = '수정 및 변환';
      seoConvertBtn.disabled = false;
    }, 2000);

  } catch (error) {
    alert(`변환 실패: ${error.message}`);
    seoConvertText.textContent = '수정 및 변환';
    seoConvertBtn.disabled = false;
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
  // 한글 파일명이 Linux 서버에서 멀티파트 파싱 에러를 일으키므로 안전한 파일명으로 전송
  formData.append('excel', selectedExcelFile, 'upload.xlsx');

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

// 사용가이드 모달
const guideBtn = document.getElementById('guide-btn');
const guideModal = document.getElementById('guide-modal');
const guideClose = document.getElementById('guide-close');
const guideOverlay = document.getElementById('guide-overlay');
const guideBody = document.getElementById('guide-body');
let guideLoaded = false;

guideBtn.addEventListener('click', () => {
  guideModal.classList.remove('hidden');
  document.body.style.overflow = 'hidden';
  if (!guideLoaded) {
    guideBody.innerHTML = marked.parse(guideMarkdown);
    guideLoaded = true;
  }
});

function closeGuide() {
  guideModal.classList.add('hidden');
  document.body.style.overflow = '';
}

guideClose.addEventListener('click', closeGuide);
guideOverlay.addEventListener('click', closeGuide);
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeGuide(); });

// 가이드 공통 스타일 (HTML/PDF 공용)
const GUIDE_STYLE = `
  body { font-family: 'Malgun Gothic', Arial, sans-serif; max-width: 860px; margin: 0 auto; padding: 32px 40px; color: #1e293b; line-height: 1.8; font-size: 14px; }
  h1 { font-size: 22px; border-bottom: 2px solid #6366f1; padding-bottom: 8px; color: #1e1b4b; }
  h2 { font-size: 17px; color: #4338ca; border-left: 4px solid #6366f1; padding-left: 10px; margin-top: 32px; }
  h3 { font-size: 15px; color: #374151; margin-top: 20px; }
  h4 { font-size: 14px; color: #6b7280; margin-top: 12px; }
  code { background: #f0f0ff; color: #4338ca; padding: 2px 6px; border-radius: 4px; font-size: 12px; font-family: Consolas, monospace; }
  pre { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px 18px; overflow-x: auto; }
  pre code { background: none; padding: 0; color: #047857; }
  table { width: 100%; border-collapse: collapse; margin: 12px 0; font-size: 13px; }
  th { background: #eef2ff; color: #3730a3; padding: 8px 12px; border: 1px solid #c7d2fe; font-weight: 600; text-align: left; }
  td { padding: 7px 12px; border: 1px solid #e2e8f0; }
  tr:nth-child(even) td { background: #f8fafc; }
  blockquote { border-left: 4px solid #6366f1; margin: 12px 0; padding: 8px 14px; background: #f5f3ff; color: #4b5563; border-radius: 0 6px 6px 0; }
  hr { border: none; border-top: 1px solid #e2e8f0; margin: 20px 0; }
  a { color: #4338ca; }
  strong { color: #111827; }
`;

// HTML 다운로드
document.getElementById('guide-html-btn').addEventListener('click', () => {
  const htmlContent = `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>VisionFlow 사용가이드</title>
  <style>${GUIDE_STYLE}</style>
</head>
<body>${marked.parse(guideMarkdown)}</body>
</html>`;
  const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'VisionFlow_사용가이드.html';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

// PDF 다운로드
document.getElementById('guide-pdf-btn').addEventListener('click', async () => {
  const btn = document.getElementById('guide-pdf-btn');
  btn.textContent = '⏳ 생성 중...';
  btn.disabled = true;

  const el = document.createElement('div');
  el.style.cssText = 'width:794px; padding:40px 50px; font-family:Malgun Gothic,Arial,sans-serif;';
  el.innerHTML = `<style>${GUIDE_STYLE}</style>` + marked.parse(guideMarkdown);

  try {
    await html2pdf().set({
      margin: 10,
      filename: 'VisionFlow_사용가이드.pdf',
      image: { type: 'jpeg', quality: 0.97 },
      html2canvas: { scale: 2, useCORS: true, logging: false },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    }).from(el).save();
  } catch (e) {
    alert('PDF 생성 실패. HTML로 다운로드 후 브라우저에서 인쇄(Ctrl+P) → PDF 저장을 이용하세요.');
  }

  btn.textContent = '⬇ PDF';
  btn.disabled = false;
});
