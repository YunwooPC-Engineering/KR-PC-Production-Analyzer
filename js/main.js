// 전역 변수
let allData = [];
let productionChart = null;

// DOM 요소
document.addEventListener('DOMContentLoaded', () => {
    const fileUpload = document.getElementById('fileUpload');
    const analyzeBtn = document.getElementById('analyzeBtn');
    const exportBtn = document.getElementById('exportBtn');
    const companySelect = document.getElementById('companySelect');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const resultSection = document.getElementById('resultSection');

    // 이벤트 리스너 등록
    analyzeBtn.addEventListener('click', analyzeFiles);
    exportBtn.addEventListener('click', exportToExcel);
    companySelect.addEventListener('change', filterByCompany);
});

// 엑셀 파일 분석
async function analyzeFiles() {
    const fileInput = document.getElementById('fileUpload');
    const files = fileInput.files;
    
    if (files.length === 0) {
        alert('파일을 선택해주세요.');
        return;
    }

    // 로딩 표시
    document.getElementById('loadingIndicator').classList.remove('d-none');
    document.getElementById('resultSection').classList.add('d-none');
    
    // 기존 데이터 초기화
    allData = [];

    try {
        // 모든 파일 처리를 병렬로 수행
        const filePromises = Array.from(files).map(file => processFile(file));
        await Promise.all(filePromises);
        
        // 데이터 정렬 (날짜순)
        allData.sort((a, b) => a.date.localeCompare(b.date));
        
        // 결과 표시
        displayResults();
    } catch (error) {
        console.error('파일 분석 중 오류 발생:', error);
        alert('파일 분석 중 오류가 발생했습니다.');
    } finally {
        // 로딩 표시 제거
        document.getElementById('loadingIndicator').classList.add('d-none');
    }
}

// 단일 엑셀 파일 처리
async function processFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // 첫 번째 시트 처리
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // 엑셀 데이터를 JSON으로 변환
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                // 파일명에서 날짜 추출 (예: "진성 생산실적_20250312.xlsx" -> "2025-03-12")
                const dateMatch = file.name.match(/(\d{8})/);
                const dateStr = dateMatch ? dateMatch[1] : "Unknown";
                const formattedDate = `${dateStr.substring(0, 4)}-${dateStr.substring(4, 6)}-${dateStr.substring(6, 8)}`;
                
                // 데이터 구조 파악 및 처리
                const processedData = parseExcelData(jsonData, formattedDate);
                
                // 전역 데이터에 추가
                allData = [...allData, ...processedData];
                
                resolve();
            } catch (error) {
                console.error('파일 처리 중 오류:', error);
                reject(error);
            }
        };
        
        reader.onerror = function(error) {
            console.error('파일 읽기 오류:', error);
            reject(error);
        };
        
        reader.readAsArrayBuffer(file);
    });
}

// 엑셀 데이터 구조 분석 및 처리
function parseExcelData(jsonData, date) {
    const processedData = [];
    let assemblyNumberIndex = -1;
    let quantityIndex = -1;
    
    // 헤더 행 찾기
    for (let i = 0; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || row.length === 0) continue;
        
        // 열 인덱스 찾기
        for (let j = 0; j < row.length; j++) {
            const cell = row[j];
            if (cell === "부재번호" || cell === "ASSEM-BLY NO." || cell && cell.toString().includes("부재")) {
                assemblyNumberIndex = j;
            }
            if (cell === "수량" || cell === "QTY" || cell && cell.toString().includes("수량")) {
                quantityIndex = j;
            }
        }
        
        // 필요한 열 인덱스를 모두 찾았으면 중단
        if (assemblyNumberIndex !== -1 && quantityIndex !== -1) {
            break;
        }
    }
    
    // 열 인덱스를 찾지 못했을 경우 기본값 설정
    if (assemblyNumberIndex === -1) assemblyNumberIndex = 1; // 일반적으로 2번째 열에 부재번호가 있다고 가정
    if (quantityIndex === -1) quantityIndex = 4; // 일반적으로 5번째 열에 수량이 있다고 가정
    
    // 데이터 행 처리
    for (let i = 0; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || row.length <= Math.max(assemblyNumberIndex, quantityIndex)) continue;
        
        const assemblyNumber = row[assemblyNumberIndex];
        const quantity = row[quantityIndex];
        
        // 부재번호와 수량이 유효한 경우만 추가
        if (assemblyNumber && quantity && !isNaN(Number(quantity))) {
            processedData.push({
                date: date,
                assemblyNumber: assemblyNumber.toString(),
                quantity: Number(quantity),
                company: document.getElementById('companySelect').value
            });
        }
    }
    
    return processedData;
}

// 결과 표시
function displayResults() {
    const resultSection = document.getElementById('resultSection');
    const tableBody = document.getElementById('resultTableBody');
    
    // 결과 섹션 표시
    resultSection.classList.remove('d-none');
    
    // 테이블 내용 초기화
    tableBody.innerHTML = '';
    
    // 날짜별, 부재번호별 데이터 집계
    const aggregatedData = aggregateData(allData);
    
    // 집계된 데이터를 테이블에 표시
    displayTable(aggregatedData);
    
    // 차트 표시
    displayChart(aggregatedData);
}

// 데이터 집계
function aggregateData(data) {
    const aggregated = {};
    
    // 현재 선택된 회사
    const selectedCompany = document.getElementById('companySelect').value;
    
    // 회사 필터링
    const filteredData = data.filter(item => item.company === selectedCompany);
    
    // 날짜별, 부재번호별 데이터 집계
    filteredData.forEach(item => {
        if (!aggregated[item.date]) {
            aggregated[item.date] = {};
        }
        
        if (!aggregated[item.date][item.assemblyNumber]) {
            aggregated[item.date][item.assemblyNumber] = 0;
        }
        
        aggregated[item.date][item.assemblyNumber] += item.quantity;
    });
    
    return aggregated;
}

// 테이블 표시
function displayTable(aggregatedData) {
    const tableBody = document.getElementById('resultTableBody');
    const dates = Object.keys(aggregatedData).sort();
    
    // 모든 부재번호 목록 생성
    const allAssemblyNumbers = new Set();
    dates.forEach(date => {
        Object.keys(aggregatedData[date]).forEach(assemblyNumber => {
            allAssemblyNumbers.add(assemblyNumber);
        });
    });
    
    // 각 부재번호별 총 생산량 계산
    const totalByAssembly = {};
    allAssemblyNumbers.forEach(assemblyNumber => {
        totalByAssembly[assemblyNumber] = 0;
        dates.forEach(date => {
            if (aggregatedData[date][assemblyNumber]) {
                totalByAssembly[assemblyNumber] += aggregatedData[date][assemblyNumber];
            }
        });
    });
    
    // 각 날짜와 부재번호에 대한 행 생성
    dates.forEach(date => {
        const dateAssemblies = Object.keys(aggregatedData[date]).sort();
        
        dateAssemblies.forEach(assemblyNumber => {
            const quantity = aggregatedData[date][assemblyNumber];
            const total = totalByAssembly[assemblyNumber];
            const progressPercent = total > 0 ? (quantity / total * 100).toFixed(1) : 0;
            
            const row = document.createElement('tr');
            
            // 날짜 셀
            const dateCell = document.createElement('td');
            dateCell.textContent = date;
            row.appendChild(dateCell);
            
            // 부재번호 셀
            const assemblyCell = document.createElement('td');
            assemblyCell.textContent = assemblyNumber;
            row.appendChild(assemblyCell);
            
            // 생산량 셀
            const quantityCell = document.createElement('td');
            quantityCell.textContent = quantity.toLocaleString();
            row.appendChild(quantityCell);
            
            // 진행률 셀
            const progressCell = document.createElement('td');
            const progressBarHTML = `
                <div class="progress">
                    <div class="progress-bar" role="progressbar" style="width: ${progressPercent}%;" 
                        aria-valuenow="${progressPercent}" aria-valuemin="0" aria-valuemax="100">
                        ${progressPercent}%
                    </div>
                </div>
            `;
            progressCell.innerHTML = progressBarHTML;
            row.appendChild(progressCell);
            
            tableBody.appendChild(row);
        });
    });
}

// 차트 표시
function displayChart(aggregatedData) {
    const dates = Object.keys(aggregatedData).sort();
    const dailyTotals = [];
    
    // 날짜별 총 생산량 계산
    dates.forEach(date => {
        let dailyTotal = 0;
        Object.keys(aggregatedData[date]).forEach(assemblyNumber => {
            dailyTotal += aggregatedData[date][assemblyNumber];
        });
        dailyTotals.push(dailyTotal);
    });
    
    // 이전 차트 제거
    if (productionChart) {
        productionChart.destroy();
    }
    
    // 차트 생성
    const ctx = document.getElementById('productionChart').getContext('2d');
    productionChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: dates,
            datasets: [{
                label: '일자별 총 생산량',
                data: dailyTotals,
                backgroundColor: 'rgba(54, 162, 235, 0.5)',
                borderColor: 'rgba(54, 162, 235, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

// 회사 필터링
function filterByCompany() {
    if (allData.length > 0) {
        displayResults();
    }
}

// Excel로 내보내기
function exportToExcel() {
    // 현재 표시된 테이블 데이터 가져오기
    const table = document.getElementById('resultTable');
    
    // 테이블 데이터를 워크시트로 변환
    const ws = XLSX.utils.table_to_sheet(table);
    
    // 워크북 생성 및 워크시트 추가
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '생산일보 분석 결과');
    
    // 현재 날짜 포맷팅
    const now = new Date();
    const dateString = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    
    // 엑셀 파일 다운로드
    XLSX.writeFile(wb, `생산일보_분석결과_${dateString}.xlsx`);
}