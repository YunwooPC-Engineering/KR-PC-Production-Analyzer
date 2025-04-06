// 전역 변수
window.allData = [];
window.filteredData = [];
window.selectedFiles = [];
window.currentSort = { column: 'date', direction: 'asc' };

// DOM이 로드되면 실행
document.addEventListener('DOMContentLoaded', function() {
    // 파일 업로드 관련 요소
    const dropZone = document.getElementById('dropZone');
    const fileUpload = document.getElementById('fileUpload');
    const selectedFilesList = document.getElementById('selectedFilesList');
    const clearFilesBtn = document.getElementById('clearFilesBtn');
    const analyzeBtn = document.getElementById('analyzeBtn');
    const exportBtn = document.getElementById('exportBtn');
    const companySelect = document.getElementById('companySelect');
    const resetFilterBtn = document.getElementById('resetFilterBtn');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const resultSection = document.getElementById('resultSection');
    const assemblyFilter = document.getElementById('assemblyFilter');
    const dateFilter = document.getElementById('dateFilter');
    const excludeItemsCheckbox = document.getElementById('excludeItemsCheckbox');
    const excludeItems = document.getElementById('excludeItems');
    const sortButtons = document.querySelectorAll('.sort-btn');

    // 드래그앤드롭 이벤트
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('active');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('active');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('active');
        handleFileSelect(e);
    });

    // 파일 선택 버튼 클릭
    dropZone.querySelector('button').addEventListener('click', () => {
        fileUpload.click();
    });

    // 파일 선택 시
    fileUpload.addEventListener('change', (e) => {
        handleFileSelect(e);
    });

    // 파일 목록 초기화
    clearFilesBtn.addEventListener('click', () => {
        selectedFiles = [];
        updateFileList();
    });

    // 분석 버튼 클릭
    analyzeBtn.addEventListener('click', () => {
        if (window.analyzeFiles) {
            window.analyzeFiles();
        } else {
            alert('분석 함수가 정의되지 않았습니다.');
        }
    });

    // 필터 관련 이벤트
    document.getElementById('assemblyFilter').addEventListener('input', applyFilters);
    document.getElementById('dateFilter').addEventListener('change', applyFilters);
    document.getElementById('excludeItemsCheckbox').addEventListener('change', applyFilters);
    document.getElementById('excludeItems').addEventListener('input', applyFilters);
    document.getElementById('resetFilterBtn').addEventListener('click', resetFilters);

    // 정렬 버튼 이벤트
    sortButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const column = btn.getAttribute('data-sort');
            toggleSort(column);
        });
    });

    // 엑셀 내보내기 버튼
    exportBtn.addEventListener('click', exportToExcel);

    // 회사 필터링
    if (companySelect) {
        companySelect.addEventListener('change', filterByCompany);
    }

    // 제외 항목 텍스트 영역 초기화
    excludeItems.disabled = !excludeItemsCheckbox.checked;
});

// 파일 선택 처리
window.handleFileSelect = function(event) {
    const files = event.target.files || event.dataTransfer.files;
    
    // 엑셀 파일만 필터링
    const excelFiles = Array.from(files).filter(file => 
        file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );

    if (excelFiles.length === 0) {
        alert('엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.');
        return;
    }

    // 선택된 파일 추가
    selectedFiles.push(...excelFiles);
    updateFileList();
};

// 파일 목록 업데이트
window.updateFileList = function() {
    const filesList = document.getElementById('selectedFilesList');
    if (!filesList) return;
    
    filesList.innerHTML = '';
    
    selectedFiles.forEach((file, index) => {
        const li = document.createElement('li');
        li.className = 'list-group-item d-flex justify-content-between align-items-center';
        li.innerHTML = `
            ${file.name}
            <button class="btn btn-sm btn-outline-danger" onclick="removeFile(${index})">
                <i class="bi bi-x"></i>
            </button>
        `;
        filesList.appendChild(li);
    });
};

// 파일 제거
window.removeFile = function(index) {
    selectedFiles.splice(index, 1);
    updateFileList();
};

// Excel로 내보내기
function exportToExcel() {
    // 현재 필터가 적용된 데이터 사용
    const currentFilteredData = filteredData;

    // 데이터가 없는 경우 처리
    if (currentFilteredData.length === 0) {
        alert('내보낼 데이터가 없습니다.');
        return;
    }

    // 워크시트에 들어갈 데이터 준비
    const wsData = [
        ['AssemblyNumber', 'Quantity', 'CompletedDate'] // 영문 필드명 사용
    ];

    // 데이터 추가 (원래 순서로 복원)
    currentFilteredData.forEach(item => {
        wsData.push([
            item.AssemblyNumber || item.assemblyNumber,
            item.Quantity || item.quantity,
            item.CompletedDate || item.date
        ]);
    });

    // 워크시트 생성
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // 열 너비 설정
    const colWidths = [
        { wch: 20 }, // AssemblyNumber
        { wch: 10 }, // Quantity
        { wch: 12 }  // CompletedDate
    ];
    ws['!cols'] = colWidths;

    // 워크북 생성 및 워크시트 추가
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '생산일보 분석 결과');

    // 현재 날짜 포맷팅
    const now = new Date();
    const dateString = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;

    // 엑셀 파일 다운로드
    XLSX.writeFile(wb, `생산일보_분석결과_${dateString}.xlsx`);
}

// 정렬 토글
function toggleSort(column) {
    // 같은 열을 클릭한 경우 정렬 방향 변경
    if (currentSort.column === column) {
        currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
    } else {
        // 다른 열을 클릭한 경우 새 열로 오름차순 정렬
        currentSort.column = column;
        currentSort.direction = 'asc';
    }
    
    // UI 아이콘 업데이트
    updateSortIcons();
    
    // 데이터 재정렬 및 표시
    sortData();
    displayFilteredData();
}

// 정렬 아이콘 업데이트
function updateSortIcons() {
    const sortButtons = document.querySelectorAll('.sort-btn');
    
    sortButtons.forEach(button => {
        const sortColumn = button.getAttribute('data-sort');
        const icon = button.querySelector('i');
        
        // 아이콘 클래스 초기화
        icon.className = '';
        
        if (sortColumn === currentSort.column) {
            // 현재 정렬 중인 열
            if (currentSort.direction === 'asc') {
                icon.className = 'bi bi-sort-alpha-down';
                if (sortColumn === 'quantity') {
                    icon.className = 'bi bi-sort-numeric-down';
                }
            } else {
                icon.className = 'bi bi-sort-alpha-up';
                if (sortColumn === 'quantity') {
                    icon.className = 'bi bi-sort-numeric-up';
                }
            }
        } else {
            // 정렬되지 않은 열
            if (sortColumn === 'quantity') {
                icon.className = 'bi bi-sort-numeric-down';
            } else {
                icon.className = 'bi bi-sort-alpha-down';
            }
        }
    });
}

// 데이터 정렬
function sortData() {
    filteredData.sort((a, b) => {
        let valueA, valueB;
        
        // 정렬할 열에 따라 비교 값 선택
        switch (currentSort.column) {
            case 'date':
                valueA = a.CompletedDate || a.date;
                valueB = b.CompletedDate || b.date;
                break;
            case 'assembly':
                valueA = a.AssemblyNumber || a.assemblyNumber;
                valueB = b.AssemblyNumber || b.assemblyNumber;
                break;
            case 'quantity':
                valueA = a.Quantity || a.quantity;
                valueB = b.Quantity || b.quantity;
                // 숫자는 문자열 비교가 아닌 숫자 비교
                return currentSort.direction === 'asc' ? valueA - valueB : valueB - valueA;
            default:
                valueA = a.CompletedDate || a.date;
                valueB = b.CompletedDate || b.date;
        }
        
        // 문자열 비교
        if (currentSort.direction === 'asc') {
            return valueA.localeCompare(valueB);
        } else {
            return valueB.localeCompare(valueA);
        }
    });
}

// 필터 초기화
function resetFilters() {
    // 필터 입력 요소 초기화
    document.getElementById('assemblyFilter').value = '';
    document.getElementById('dateFilter').value = 'all';
    document.getElementById('excludeItemsCheckbox').checked = false;
    document.getElementById('excludeItems').value = '';
    document.getElementById('excludeItems').disabled = true;
    
    // 필터 적용하여 모든 데이터 표시
    filteredData = [...allData];
    
    // 정렬 적용
    sortData();
    
    // 결과 표시
    displayFilteredData();
    
    console.log('필터가 초기화되었습니다.');
}

// 필터 적용
function applyFilters() {
    const assemblyFilterValue = document.getElementById('assemblyFilter').value.toLowerCase();
    const dateFilterValue = document.getElementById('dateFilter').value;
    const excludeItemsCheckbox = document.getElementById('excludeItemsCheckbox');
    const excludeItemsInput = document.getElementById('excludeItems');
    
    // 제외할 부재번호 목록 생성
    let excludedAssemblies = [];
    if (excludeItemsCheckbox.checked) {
        excludedAssemblies = excludeItemsInput.value
            .split(',')
            .map(item => item.trim())
            .filter(item => item !== '');
        
        console.log('제외할 부재번호 패턴:', excludedAssemblies);
    }
    
    // 필터링
    filteredData = allData.filter(item => {
        // 부재번호 필터
        const matchesAssembly = !assemblyFilterValue || 
            item.assemblyNumber.toLowerCase().includes(assemblyFilterValue);
            
        // 날짜 필터
        const matchesDate = dateFilterValue === 'all' || item.date === dateFilterValue;
        
        // 제외 부재번호 필터 (부분 일치 방식)
        const isExcluded = excludedAssemblies.length > 0 && 
            excludedAssemblies.some(pattern => item.assemblyNumber.includes(pattern));
            
        return matchesAssembly && matchesDate && !isExcluded;
    });
    
    console.log('필터 적용 후 데이터 수:', filteredData.length);
    
    // 정렬 적용
    sortData();
    
    // 결과 표시
    displayFilteredData();
}

// 회사 필터링
function filterByCompany() {
    if (allData.length > 0) {
        applyFilters();
    }
}

// 필터링된 데이터 표시
window.displayFilteredData = function() {
    const tbody = document.getElementById('resultTableBody');
    if (!tbody) return;
    
    tbody.innerHTML = '';

    filteredData.forEach(item => {
        const tr = document.createElement('tr');
        const date = item.CompletedDate || item.date;
        const assemblyNumber = item.AssemblyNumber || item.assemblyNumber;
        const quantity = item.Quantity || item.quantity;
        
        tr.innerHTML = `
            <td>${formatDate(date)}</td>
            <td>${assemblyNumber}</td>
            <td>${quantity.toLocaleString()}</td>
        `;
        tbody.appendChild(tr);
    });

    // 합계 업데이트
    updateSummary();
};

// 요약 정보 업데이트
window.updateSummary = function() {
    // 총 부재 유형 수
    const totalAssemblyTypes = document.getElementById('totalAssemblyTypes');
    if (totalAssemblyTypes) {
        totalAssemblyTypes.textContent = 
            [...new Set(filteredData.map(item => item.AssemblyNumber || item.assemblyNumber))].length.toLocaleString();
    }
    
    // 총 생산량
    const totalProductionQuantity = document.getElementById('totalProductionQuantity');
    if (totalProductionQuantity) {
        totalProductionQuantity.textContent = 
            filteredData.reduce((sum, item) => sum + (item.Quantity || item.quantity), 0).toLocaleString();
    }
    
    // 총 부재 수
    const totalAssemblyCount = document.getElementById('totalAssemblyCount');
    if (totalAssemblyCount) {
        totalAssemblyCount.textContent = 
            `${filteredData.length.toLocaleString()}개 부재`;
    }
    
    // 총량
    const totalQuantity = document.getElementById('totalQuantity');
    if (totalQuantity) {
        totalQuantity.textContent = 
            filteredData.reduce((sum, item) => sum + (item.Quantity || item.quantity), 0).toLocaleString();
    }
};

// 회사별 파일 분석 함수
window.analyzeFiles = async function() {
    const files = selectedFiles;
    
    if (files.length === 0) {
        alert('파일을 선택해주세요.');
        return;
    }

    // 로딩 표시
    document.getElementById('loadingIndicator').classList.remove('d-none');
    document.getElementById('resultSection').classList.add('d-none');
    
    // 기존 데이터 초기화
    allData = [];
    filteredData = [];

    try {
        const company = document.getElementById('companySelect').value;
        console.log('선택된 회사:', company);
        
        // 회사별 분석 처리
        switch (company) {
            case 'jinsungpc':
                // 모든 파일 처리를 병렬로 수행
                const filePromises = Array.from(files).map(file => processFile(file));
                const results = await Promise.all(filePromises);
                
                // 모든 결과를 allData에 추가
                results.forEach(result => {
                    if (result && Array.isArray(result)) {
                        allData = [...allData, ...result];
                    }
                });
                break;
                
            case 'isue':
                // 이수이앤씨는 자체 분석기 사용
                if (typeof analyzeIsueFiles === 'function') {
                    allData = await analyzeIsueFiles();
                } else {
                    throw new Error('이수이앤씨 분석기가 로드되지 않았습니다.');
                }
                break;
                
            case 'jisan':
                // 지산은 자체 분석기 사용
                if (typeof analyzeJisanFiles === 'function') {
                    allData = await analyzeJisanFiles();
                } else {
                    throw new Error('지산 분석기가 로드되지 않았습니다.');
                }
                break;
                
            default:
                throw new Error('지원하지 않는 회사입니다.');
        }
        
        console.log('처리된 전체 데이터:', allData.length, '건');
        
        // 데이터 정렬 (날짜순)
        allData.sort((a, b) => a.date.localeCompare(b.date));
        
        // 같은 날짜의 중복 부재번호 제거 (최신 파일의 데이터 유지)
        const uniqueItems = new Map();
        for (let i = allData.length - 1; i >= 0; i--) {
            const item = allData[i];
            const key = `${item.date}-${item.assemblyNumber}`;
            uniqueItems.set(key, item);
        }
        allData = Array.from(uniqueItems.values());
        
        // 필터링된 데이터 설정
        filteredData = [...allData];
        
        // 날짜 필터 옵션 업데이트
        updateDateFilter();
        
        // 결과 표시
        displayFilteredData();
    } catch (error) {
        console.error('파일 분석 중 오류 발생:', error);
        alert('파일 분석 중 오류가 발생했습니다: ' + error.message);
    } finally {
        // 로딩 표시 제거
        document.getElementById('loadingIndicator').classList.add('d-none');
    }
}

// 날짜 필터 옵션 업데이트
window.updateDateFilter = function() {
    const dateFilter = document.getElementById('dateFilter');
    if (!dateFilter) return;
    
    const dates = [...new Set(allData.map(item => item.CompletedDate || item.date))].sort();
    
    dateFilter.innerHTML = '<option value="all">전체</option>' + 
        dates.map(date => `<option value="${date}">${formatDate(date)}</option>`).join('');
};

// 날짜 형식 변환
window.formatDate = function(dateStr) {
    if (!dateStr) return '';
    return `${dateStr.substring(0, 4)}년 ${dateStr.substring(4, 6)}월 ${dateStr.substring(6, 8)}일`;
};

// 단일 엑셀 파일 처리
async function processFile(file) {
    return new Promise((resolve, reject) => {
        // 파일 수정 날짜 추출
        const fileModifiedDate = new Date(file.lastModified);
        const modifiedDate = formatDateFromDate(fileModifiedDate);
        
        // 파일명에서 날짜 추출 (yyMMdd, yyyyMMdd, yyyy-MM-dd)
        let filenameDate = extractDateFromText(file.name);
        
        console.log(`파일 [${file.name}] 처리 시작`);
        console.log(`- 파일 수정일: ${modifiedDate}`);
        console.log(`- 파일명에서 추출한 날짜: ${filenameDate || '없음'}`);
        
        // 파일 데이터 처리 로직 (읽고, 파싱하고, 저장하는 과정)
        const reader = new FileReader();
            
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // 시트명에서 날짜 추출 시도
                let sheetNameDate = null;
                for (const sheetName of workbook.SheetNames) {
                    const extractedDate = extractDateFromText(sheetName);
                    if (extractedDate) {
                        sheetNameDate = extractedDate;
                        console.log(`- 시트명 ${sheetName}에서 추출한 날짜: ${sheetNameDate}`);
                        break;
                    }
                }
                
                // 첫 번째 시트 처리
                const firstSheetName = workbook.SheetNames[0];
                console.log('- 첫 번째 시트 이름:', firstSheetName);
                const worksheet = workbook.Sheets[firstSheetName];
                
                // 엑셀 데이터를 JSON으로 변환 (헤더 포함)
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1,
                    defval: '',
                    raw: false
                });
                console.log('- 엑셀 데이터 변환 완료:', jsonData.length, '행');
                
                // 문서 내용에서 날짜 추출 시도
                const documentDate = extractDateFromDocument(jsonData);
                console.log(`- 문서 내용에서 추출한 날짜: ${documentDate || '없음'}`);
                
                // 가장 최신 날짜 선택
                const availableDates = [filenameDate, modifiedDate, sheetNameDate, documentDate]
                    .filter(date => date); // null/undefined 제거
                
                let finalDate;
                if (availableDates.length > 0) {
                    // 날짜 비교하여 가장 최신 날짜 선택
                    finalDate = selectMostRecentDate(availableDates);
                } else {
                    // 날짜를 하나도 찾지 못한 경우 현재 날짜 사용
                    const now = new Date();
                    finalDate = formatDateFromDate(now);
                }
                
                console.log(`- 최종 선택된 날짜: ${finalDate}`);
                
                // 공장 이름 확인
                const factoryName = determineFactory(file.name);
                console.log(`- 결정된 공장 이름: ${factoryName}`);
                
                // 데이터 파싱
                const processedData = parseFactoryData(jsonData, finalDate, factoryName);
                console.log(`- 처리된 데이터 수: ${processedData.length}`);
                
                // 처리 결과 반환
                resolve(processedData);
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

// Date 객체에서 YYYYMMDD 형식의 문자열로 변환
function formatDateFromDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}${month}${day}`;
}

// 텍스트에서 날짜 추출 (파일명, 시트명 등에 사용)
function extractDateFromText(text) {
    if (!text) return null;
    
    // YYYYMMDD, YYYY-MM-DD, YYYY_MM_DD 패턴
    const fullDatePattern = /(\d{4})[-_]?(\d{2})[-_]?(\d{2})/;
    // MMDD, MM-DD 패턴
    const shortDatePattern = /(\d{2})[-_]?(\d{2})/;
    // 년월일 패턴 (2023년 3월 15일, 2023.3.15, 2023-3-15)
    const koreanDatePattern = /(\d{4})[-년\.\-]?\s*(\d{1,2})[-월\.\-]?\s*(\d{1,2})[-일]?/;
    // 2자리 연도 패턴 (23년 3월 15일, 23.3.15)
    const shortYearPattern = /(\d{2})[-년\.\-]?\s*(\d{1,2})[-월\.\-]?\s*(\d{1,2})[-일]?/;
    
    // 현재 날짜 가져오기
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1;
    const currentDay = now.getDate();
    
    // 날짜 유효성 검사 함수
    const isValidDate = (year, month, day) => {
        // 기본 유효성 검사
        if (month < 1 || month > 12 || day < 1 || day > 31) return false;
        
        // 월별 일수 검사
        const daysInMonth = new Date(year, month, 0).getDate();
        if (day > daysInMonth) return false;
        
        // 미래 날짜 확인 (현재보다 2년 이상 미래는 오류로 간주)
        if (year > currentYear + 2) {
            console.warn(`유효하지 않은 미래 연도: ${year}, 현재 연도로 조정합니다.`);
            return false;
        }
        
        // 올해보다 미래인 경우 현재 이후의 날짜인지 확인
        if (year === currentYear && 
            (month > currentMonth || (month === currentMonth && day > currentDay))) {
            console.warn(`미래 날짜 감지: ${year}-${month}-${day}, 년도를 작년으로 조정합니다.`);
            return false;
        }
        
        return true;
    };
    
    // 날짜 조정 함수 (미래 날짜 처리)
    const adjustDate = (year, month, day) => {
        // 2자리 연도를 4자리로 변환 시 미래 날짜 처리
        if (year.toString().length === 2) {
            // 2자리 연도를 4자리로 변환 (20xx 형식)
            let fullYear = parseInt(`20${year}`);
            
            // 미래 연도인 경우 100년 전으로 조정 (2025 -> 1925)
            if (fullYear > currentYear + 2) {
                fullYear = parseInt(`19${year}`);
            }
            
            year = fullYear;
        }
        
        // 미래 날짜 조정 (올해보다 미래인 경우)
        if (year > currentYear + 2) {
            console.warn(`미래 연도 조정: ${year} -> ${currentYear}`);
            year = currentYear;
        } else if (year === currentYear && 
                  (month > currentMonth || (month === currentMonth && day > currentDay))) {
            // 올해인데 현재 날짜보다 미래인 경우 작년으로 조정
            console.warn(`현재 이후의 날짜 감지: ${year}-${month}-${day}, 년도를 조정합니다.`);
            year = currentYear - 1;
        }
        
        return { year, month, day };
    };
    
    // YYYYMMDD 패턴 확인
    let match = text.match(fullDatePattern);
    if (match) {
        let year = parseInt(match[1]);
        let month = parseInt(match[2]);
        let day = parseInt(match[3]);
        
        const adjusted = adjustDate(year, month, day);
        if (isValidDate(adjusted.year, adjusted.month, adjusted.day)) {
            return `${adjusted.year}${String(adjusted.month).padStart(2, '0')}${String(adjusted.day).padStart(2, '0')}`;
        }
    }
    
    // 한국어 날짜 패턴 확인
    match = text.match(koreanDatePattern);
    if (match) {
        let year = parseInt(match[1]);
        let month = parseInt(match[2]);
        let day = parseInt(match[3]);
        
        const adjusted = adjustDate(year, month, day);
        if (isValidDate(adjusted.year, adjusted.month, adjusted.day)) {
            return `${adjusted.year}${String(adjusted.month).padStart(2, '0')}${String(adjusted.day).padStart(2, '0')}`;
        }
    }
    
    // 2자리 연도 패턴 확인
    match = text.match(shortYearPattern);
    if (match) {
        let year = parseInt(match[1]);
        let month = parseInt(match[2]);
        let day = parseInt(match[3]);
        
        // 2자리 연도는 항상 20XX로 처리 후 유효성 검사
        const adjusted = adjustDate(year, month, day);
        if (isValidDate(adjusted.year, adjusted.month, adjusted.day)) {
            return `${adjusted.year}${String(adjusted.month).padStart(2, '0')}${String(adjusted.day).padStart(2, '0')}`;
        }
    }
    
    // MMDD 패턴 확인
    match = text.match(shortDatePattern);
    if (match) {
        let month = parseInt(match[1]);
        let day = parseInt(match[2]);
        
        // 현재 월/일보다 미래인 경우 작년으로 처리
        let year = currentYear;
        if (month > currentMonth || (month === currentMonth && day > currentDay)) {
            year = currentYear - 1;
        }
        
        if (isValidDate(year, month, day)) {
            return `${year}${String(month).padStart(2, '0')}${String(day).padStart(2, '0')}`;
        }
    }
    
    return null;
}

// 문서 내용에서 날짜 추출
function extractDateFromDocument(jsonData) {
    if (!jsonData || jsonData.length === 0) return null;
    
    // 처음 10행 검색
    for (let i = 0; i < Math.min(10, jsonData.length); i++) {
        const row = jsonData[i];
        if (!row) continue;
        
        // 각 셀의 값을 문자열로 변환하여 검사
        for (let j = 0; j < row.length; j++) {
            const cellValue = String(row[j] || '');
            const extractedDate = extractDateFromText(cellValue);
            
            if (extractedDate) {
                console.log(`문서 내용 [${i}행, ${j}열]에서 날짜 발견: ${extractedDate}, 값: ${cellValue}`);
                return extractedDate;
            }
        }
    }
    
    return null;
}

// 여러 날짜 중 가장 최신 날짜 선택
function selectMostRecentDate(dates) {
    return dates.reduce((latest, current) => {
        if (!latest) return current;
        
        // YYYYMMDD 형식 문자열을 Date 객체로 변환하여 비교
        const latestDate = new Date(
            latest.substring(0, 4), 
            parseInt(latest.substring(4, 6)) - 1, 
            latest.substring(6, 8)
        );
        
        const currentDate = new Date(
            current.substring(0, 4), 
            parseInt(current.substring(4, 6)) - 1, 
            current.substring(6, 8)
        );
        
        return currentDate > latestDate ? current : latest;
    }, null);
}

// 파일명으로 공장 종류 판단하는 함수
function determineFactory(fileName) {
    fileName = fileName.toLowerCase();
    
    if (fileName.includes('진성') || fileName.includes('jinsungpc')) {
        return 'jinsungpc';
    } else if (fileName.includes('여주') || fileName.includes('yeoju')) {
        return 'esue_yeoju';
    } else if (fileName.includes('이수') || fileName.includes('isue') || fileName.includes('음성') || fileName.includes('eumseong')) {
        return 'isue_eumseong';
    } else if (fileName.includes('지산') || fileName.includes('jisan')) {
        return 'jisan';
    } else if (fileName.includes('나라') || fileName.includes('narapc')) {
        return 'narapc';
    } else if (fileName.includes('부산bgf') || fileName.includes('busanbgf')) {
        return 'busanbgf';
    } else {
        // 기본값은 진성피씨로 설정
        return 'jinsungpc';
    }
}

// 공장별 데이터 파싱 함수
function parseFactoryData(jsonData, fileDate, factoryName) {
    console.log(`${factoryName} 파싱 시작 (파일 날짜: ${fileDate})`);
    
    // 공장별 파싱 로직 적용
    switch (factoryName) {
        case 'jinsungpc':
            return parseJinsungPCData(jsonData, fileDate);
        case 'isue_eumseong':
            if (window.parseIsueData) {
                return window.parseIsueData(jsonData, fileDate);
            }
            break;
        case 'esue_yeoju':
            if (window.EsueYeojuDataParser) {
                const parser = new window.EsueYeojuDataParser();
                return parser.parseSheetData(jsonData, fileDate);
            }
            break;
        case 'jisan':
            if (window.parseJisanData) {
                return window.parseJisanData(jsonData, fileDate);
            }
            break;
        case 'narapc':
            if (window.parseNaraPCData) {
                return window.parseNaraPCData(jsonData, fileDate);
            }
            break;
        case 'busanbgf':
            if (window.parseBusanBGFData) {
                return window.parseBusanBGFData(jsonData, fileDate);
            }
            break;
        default:
            console.warn(`지원되지 않는 공장 형식: ${factoryName}. 빈 배열 반환.`);
            return [];
    }
    
    console.warn(`${factoryName} 파서를 찾을 수 없습니다. 빈 배열 반환.`);
    return [];
}