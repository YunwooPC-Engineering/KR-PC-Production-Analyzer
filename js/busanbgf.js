// 부산BGF 물류 생산일보 분석
async function analyzeFiles() {
    if (selectedFiles.length === 0) {
        alert('분석할 파일을 선택해주세요.');
        return;
    }

    // 로딩 표시
    document.getElementById('loadingIndicator').classList.remove('d-none');
    document.getElementById('resultSection').classList.add('d-none');

    try {
        // 파일들을 병렬로 처리
        const filePromises = selectedFiles.map(file => processFile(file));
        const results = await Promise.all(filePromises);

        // 모든 데이터를 하나의 배열로 합치기
        allData = results.flat();

        // 날짜별로 정렬
        allData.sort((a, b) => a.date.localeCompare(b.date));

        // 날짜 필터 옵션 업데이트
        updateDateFilterOptions();

        // 필터 적용 및 결과 표시
        applyFilters();
        displayFilteredData();

        // 결과 섹션 표시
        document.getElementById('resultSection').classList.remove('d-none');
    } catch (error) {
        console.error('파일 분석 중 오류 발생:', error);
        alert('파일 분석 중 오류가 발생했습니다.');
    } finally {
        // 로딩 표시 숨기기
        document.getElementById('loadingIndicator').classList.add('d-none');
    }
}

// 파일 처리
async function processFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 'A' });

                // 파일명에서 날짜 추출 (YYYYMMDD 형식)
                const dateMatch = file.name.match(/(\d{8})/);
                const date = dateMatch ? dateMatch[1] : '';

                // 데이터 파싱
                const parsedData = parseBusanBGFData(jsonData, date);
                resolve(parsedData);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// 부산BGF 데이터 파싱
function parseBusanBGFData(jsonData, date) {
    const parsedData = [];
    
    // 부재번호와 수량이 있는 열 찾기
    let assemblyCol = null;
    let quantityCol = null;
    
    // 헤더 행에서 열 찾기
    for (let i = 0; i < jsonData.length; i++) {
        const row = jsonData[i];
        for (const key in row) {
            const value = row[key];
            if (typeof value === 'string') {
                if (value.includes('부재번호') || value.includes('Assembly')) {
                    assemblyCol = key;
                } else if (value.includes('수량') || value.includes('Quantity')) {
                    quantityCol = key;
                }
            }
        }
        if (assemblyCol && quantityCol) break;
    }
    
    // 데이터 행 처리
    for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        const assemblyNumber = row[assemblyCol];
        const quantity = row[quantityCol];
        
        // 부재번호와 수량이 있는 행만 처리
        if (assemblyNumber && quantity && 
            !assemblyNumber.toString().toLowerCase().includes('소계') &&
            !assemblyNumber.toString().toLowerCase().includes('합계') &&
            !assemblyNumber.toString().toLowerCase().includes('total') &&
            !assemblyNumber.toString().toLowerCase().includes('subtotal')) {
            
            parsedData.push({
                date: date,
                assemblyNumber: assemblyNumber.toString().trim(),
                quantity: parseInt(quantity) || 0
            });
        }
    }
    
    return parsedData;
}

// 날짜 필터 옵션 업데이트
function updateDateFilterOptions() {
    const dateFilter = document.getElementById('dateFilter');
    const uniqueDates = [...new Set(allData.map(item => item.date))].sort();
    
    // 기존 옵션 초기화
    dateFilter.innerHTML = '<option value="all">전체</option>';
    
    // 날짜 옵션 추가
    uniqueDates.forEach(date => {
        const option = document.createElement('option');
        option.value = date;
        option.textContent = date;
        dateFilter.appendChild(option);
    });
}

// 필터링된 데이터 표시
function displayFilteredData() {
    const tbody = document.querySelector('#resultTable tbody');
    tbody.innerHTML = '';
    
    // 소계 행을 제외한 데이터만 표시
    filteredData.forEach(item => {
        if (!item.assemblyNumber.toLowerCase().includes('소계') &&
            !item.assemblyNumber.toLowerCase().includes('합계') &&
            !item.assemblyNumber.toLowerCase().includes('total') &&
            !item.assemblyNumber.toLowerCase().includes('subtotal')) {
            addTableRow(item);
        }
    });
}

// 테이블 행 추가
function addTableRow(item) {
    const tbody = document.querySelector('#resultTable tbody');
    const row = document.createElement('tr');
    
    row.innerHTML = `
        <td>${item.date}</td>
        <td>${item.assemblyNumber}</td>
        <td>${item.quantity}</td>
    `;
    
    tbody.appendChild(row);
} 