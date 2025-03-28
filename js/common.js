// 전역 변수
let allData = [];
let filteredData = [];
let currentSort = { column: null, direction: 'asc' };
let selectedFiles = [];

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
        handleFiles(e.dataTransfer.files);
    });

    // 파일 선택 버튼 클릭
    dropZone.querySelector('button').addEventListener('click', () => {
        fileUpload.click();
    });

    // 파일 선택 시
    fileUpload.addEventListener('change', (e) => {
        handleFiles(e.target.files);
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
    companySelect.addEventListener('change', filterByCompany);

    // 제외 항목 텍스트 영역 초기화
    excludeItems.disabled = !excludeItemsCheckbox.checked;
});

// 파일 처리 함수
function handleFiles(files) {
    // .xlsx 파일만 필터링
    const xlsxFiles = Array.from(files).filter(file => file.name.endsWith('.xlsx'));
    
    if (xlsxFiles.length === 0) {
        alert('엑셀(.xlsx) 파일만 선택 가능합니다.');
        return;
    }

    // 선택된 파일 목록 업데이트
    selectedFiles = [...selectedFiles, ...xlsxFiles];
    updateFileList();
}

// 파일 목록 업데이트
function updateFileList() {
    const fileList = document.getElementById('fileList');
    const selectedFilesList = document.getElementById('selectedFilesList');
    
    if (selectedFiles.length > 0) {
        fileList.classList.remove('d-none');
        selectedFilesList.innerHTML = selectedFiles.map((file, index) => `
            <li class="list-group-item d-flex justify-content-between align-items-center">
                ${file.name}
                <button class="btn btn-sm btn-outline-danger" onclick="removeFile(${index})">
                    <i class="bi bi-x"></i>
                </button>
            </li>
        `).join('');
    } else {
        fileList.classList.add('d-none');
    }
}

// 파일 제거
function removeFile(index) {
    selectedFiles.splice(index, 1);
    updateFileList();
}

// Excel로 내보내기
function exportToExcel() {
    // 현재 필터가 적용된 데이터 가져오기
    const currentFilteredData = getCurrentFilteredData();

    // 데이터가 없는 경우 처리
    if (currentFilteredData.length === 0) {
        alert('내보낼 데이터가 없습니다.');
        return;
    }

    // 워크시트에 들어갈 데이터 준비
    const wsData = [
        ['AssemblyNumber', 'Quantity', 'CompletedDate'] // 헤더 행
    ];

    // 필터링된 데이터를 배열에 추가
    currentFilteredData.forEach(item => {
        wsData.push([
            item.assemblyNumber,
            item.quantity,
            item.date
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

// 현재 필터가 적용된 데이터 가져오기
function getCurrentFilteredData() {
    const assemblyFilterValue = document.getElementById('assemblyFilter').value.toLowerCase();
    const dateFilterValue = document.getElementById('dateFilter').value;
    const excludeItemsCheckbox = document.getElementById('excludeItemsCheckbox');
    
    // 제외할 부재번호 목록
    let excludedAssemblies = [];
    if (excludeItemsCheckbox.checked) {
        const excludeItemsValue = document.getElementById('excludeItems').value;
        excludedAssemblies = excludeItemsValue.split(',')
            .map(item => item.trim().toLowerCase())
            .filter(item => item !== '');
    }
    
    // 필터링된 데이터 반환
    return allData.filter(item => {
        // 부재번호 필터
        const matchesAssembly = !assemblyFilterValue || 
            item.assemblyNumber.toLowerCase().includes(assemblyFilterValue);
        
        // 날짜 필터
        const matchesDate = dateFilterValue === 'all' || 
            item.date === dateFilterValue;
        
        // 제외 항목 필터
        const isNotExcluded = excludedAssemblies.length === 0 || 
            !excludedAssemblies.some(excluded => 
                item.assemblyNumber.toLowerCase().includes(excluded)
            );
        
        return matchesAssembly && matchesDate && isNotExcluded;
    });
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
                valueA = a.date;
                valueB = b.date;
                break;
            case 'assembly':
                valueA = a.assemblyNumber;
                valueB = b.assemblyNumber;
                break;
            case 'quantity':
                valueA = a.quantity;
                valueB = b.quantity;
                // 숫자는 문자열 비교가 아닌 숫자 비교
                return currentSort.direction === 'asc' ? valueA - valueB : valueB - valueA;
            default:
                valueA = a.date;
                valueB = b.date;
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
    document.getElementById('assemblyFilter').value = '';
    document.getElementById('dateFilter').value = 'all';
    document.getElementById('excludeItemsCheckbox').checked = false;
    document.getElementById('excludeItems').value = '';
    document.getElementById('excludeItems').disabled = true;
    applyFilters();
}

// 필터 적용
function applyFilters() {
    const assemblyFilterValue = document.getElementById('assemblyFilter').value.toLowerCase();
    const dateFilterValue = document.getElementById('dateFilter').value;
    const excludeItemsCheckbox = document.getElementById('excludeItemsCheckbox');
    
    // 제외할 부재번호 목록
    let excludedAssemblies = [];
    if (excludeItemsCheckbox.checked) {
        const excludeItemsValue = document.getElementById('excludeItems').value;
        excludedAssemblies = excludeItemsValue.split(',')
            .map(item => item.trim().toLowerCase())
            .filter(item => item !== '');
    }
    
    // 필터링된 데이터 업데이트
    filteredData = allData.filter(item => {
        // 부재번호 필터
        const matchesAssembly = !assemblyFilterValue || 
            item.assemblyNumber.toLowerCase().includes(assemblyFilterValue);
        
        // 날짜 필터
        const matchesDate = dateFilterValue === 'all' || 
            item.date === dateFilterValue;
        
        // 제외 항목 필터
        const isNotExcluded = excludedAssemblies.length === 0 || 
            !excludedAssemblies.some(excluded => 
                item.assemblyNumber.toLowerCase().includes(excluded)
            );
        
        return matchesAssembly && matchesDate && isNotExcluded;
    });
    
    // 정렬 및 표시
    sortData();
    displayFilteredData();
}

// 회사 필터링
function filterByCompany() {
    if (allData.length > 0) {
        applyFilters();
    }
}

// 필터링된 데이터 표시
function displayFilteredData() {
    const resultSection = document.getElementById('resultSection');
    const tableBody = document.getElementById('resultTableBody');
    
    // 결과 섹션 표시
    resultSection.classList.remove('d-none');
    
    // 테이블 내용 초기화
    tableBody.innerHTML = '';
    
    // 테이블 생성
    if (filteredData.length > 0) {
        // 날짜별 그룹화
        const dateGroups = {};
        filteredData.forEach(item => {
            if (!dateGroups[item.date]) {
                dateGroups[item.date] = [];
            }
            dateGroups[item.date].push(item);
        });
        
        // 날짜별 데이터 표시
        Object.keys(dateGroups).sort().forEach(date => {
            // 날짜별 데이터
            const dateData = dateGroups[date];
            
            // 각 부재번호별 행 추가
            dateData.forEach(item => {
                const row = document.createElement('tr');
                
                // 날짜 셀
                const dateCell = document.createElement('td');
                dateCell.textContent = item.date;
                row.appendChild(dateCell);
                
                // 부재번호 셀
                const assemblyCell = document.createElement('td');
                assemblyCell.textContent = item.assemblyNumber;
                row.appendChild(assemblyCell);
                
                // 생산량 셀
                const quantityCell = document.createElement('td');
                quantityCell.textContent = item.quantity.toLocaleString();
                row.appendChild(quantityCell);
                
                tableBody.appendChild(row);
            });
            
            // 날짜별 소계 행 추가
            const totalQuantity = dateData.reduce((sum, item) => sum + item.quantity, 0);
            const uniqueAssemblyCount = new Set(dateData.map(item => item.assemblyNumber)).size;
            
            const subtotalRow = document.createElement('tr');
            subtotalRow.className = 'table-secondary';
            
            const subtotalDateCell = document.createElement('td');
            subtotalDateCell.textContent = `${date} 소계`;
            subtotalRow.appendChild(subtotalDateCell);
            
            const subtotalAssemblyCell = document.createElement('td');
            subtotalAssemblyCell.textContent = `${uniqueAssemblyCount}개 부재`;
            subtotalRow.appendChild(subtotalAssemblyCell);
            
            const subtotalQuantityCell = document.createElement('td');
            subtotalQuantityCell.textContent = totalQuantity.toLocaleString();
            subtotalRow.appendChild(subtotalQuantityCell);
            
            tableBody.appendChild(subtotalRow);
        });
        
        // 총계 업데이트
        const totalQuantity = filteredData.reduce((sum, item) => sum + item.quantity, 0);
        const uniqueAssemblyNumbers = new Set(filteredData.map(item => item.assemblyNumber));
        
        document.getElementById('totalAssemblyCount').textContent = `${uniqueAssemblyNumbers.size}개 부재`;
        document.getElementById('totalQuantity').textContent = totalQuantity.toLocaleString();
    } else {
        // 데이터가 없는 경우
        const emptyRow = document.createElement('tr');
        const emptyCell = document.createElement('td');
        emptyCell.colSpan = 3;
        emptyCell.textContent = '데이터가 없습니다.';
        emptyCell.className = 'text-center';
        emptyRow.appendChild(emptyCell);
        tableBody.appendChild(emptyRow);
        
        // 합계 초기화
        document.getElementById('totalAssemblyCount').textContent = '0개 부재';
        document.getElementById('totalQuantity').textContent = '0';
    }
}

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
                const filePromises = Array.from(files).map(file => processFile(file, parseJinsungPCData));
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

// 단일 엑셀 파일 처리
async function processFile(file, parseFunction) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // 첫 번째 시트 처리
                const firstSheetName = workbook.SheetNames[0];
                console.log('시트 이름:', firstSheetName);
                const worksheet = workbook.Sheets[firstSheetName];
                
                // 엑셀 데이터를 JSON으로 변환 (헤더 포함)
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1,
                    defval: '',
                    raw: false
                });
                console.log('엑셀 데이터 변환 완료:', jsonData.length, '행');
                
                // 파일명에서 날짜 추출 (예: "진성 생산실적_20250312.xlsx" -> "2025-03-12")
                const dateMatch = file.name.match(/(\d{8})/);
                if (!dateMatch) {
                    console.error('파일명에서 날짜를 찾을 수 없습니다:', file.name);
                    resolve([]);
                    return;
                }
                
                const dateStr = dateMatch[1];
                const formattedDate = `${dateStr.substring(0, 4)}-${dateStr.substring(4, 6)}-${dateStr.substring(6, 8)}`;
                
                // 데이터 파싱
                console.log(`파일 처리 시작: ${file.name}`);
                const processedData = parseFunction(jsonData, formattedDate);
                console.log(`파일 처리 완료: ${file.name}, 처리된 데이터 수: ${processedData.length}`);
                
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