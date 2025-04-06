// 나라피씨 생산일보 분석
class NaraPCAnalyzer {
    constructor() {
        this.factoryName = '나라피씨';
        this.excludedKeywords = ['소계', '합계', 'total', 'subtotal'];
    }

    // 날짜 형식 변환 (YYYYMMDD -> YYYY-MM-DD)
    formatDate(dateString) {
        if (dateString.length === 8) {
            return `${dateString.substring(0, 4)}-${dateString.substring(4, 6)}-${dateString.substring(6)}`;
        }
        return dateString;
    }

    // 파일 이름에서 날짜 추출
    extractDateFromFilename(filename) {
        // 패턴: YYYYMMDD 또는 MMDD
        const fullDatePattern = /(\d{4})(\d{2})(\d{2})/;
        const shortDatePattern = /(\d{2})(\d{2})/;
        
        let match = filename.match(fullDatePattern);
        if (match) {
            return match[0]; // YYYYMMDD 형식
        }
        
        match = filename.match(shortDatePattern);
        if (match) {
            const currentYear = new Date().getFullYear();
            const month = match[1];
            const day = match[2];
            return `${currentYear}${month}${day}`;
        }
        
        // 날짜를 찾지 못한 경우 현재 날짜 사용
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        return `${year}${month}${day}`;
    }
    
    // 시트 이름에서 날짜 추출
    extractDateFromSheetName(sheetName) {
        // 시트 이름이 "MMDD" 형식이면 날짜로 변환
        const datePattern = /^(\d{2})(\d{2})$/;
        const match = String(sheetName).match(datePattern);
        
        if (match) {
            const currentYear = new Date().getFullYear();
            const month = match[1];
            const day = match[2];
            return `${currentYear}${month}${day}`;
        }
        
        return null;
    }

    // 제외할 키워드 체크
    isExcludedRow(assemblyNumber) {
        if (!assemblyNumber) return true;
        
        const lowerStr = String(assemblyNumber).toLowerCase();
        return this.excludedKeywords.some(keyword => lowerStr.includes(keyword));
    }

    parseFactoryData(workbook, fileDate) {
        let allParsedData = [];
        
        // 모든 시트를 처리
        for (const sheetName of workbook.SheetNames) {
            // 제외할 시트 이름 (사용금지, 미사용 등의 문자가 포함된 시트는 건너뜀)
            if (sheetName.includes('사용금지') || 
                sheetName.includes('미사용') || 
                sheetName.includes('양식') || 
                sheetName.includes('폐기')) {
                console.log(`건너뛴 시트: ${sheetName} (제외 시트)`);
                continue;
            }
            
            console.log(`시트 처리 중: ${sheetName}`);
            
            // 시트 이름에서 날짜 추출 시도
            const sheetDate = this.extractDateFromSheetName(sheetName);
            
            // 최종 사용할 날짜 결정 (시트 날짜 > 파일 날짜)
            const finalDate = sheetDate || fileDate;
            console.log(`시트 ${sheetName}의 사용 날짜: ${finalDate}`);
            
            const worksheet = workbook.Sheets[sheetName];
            
            // 시트의 병합 셀 정보 및 범위 확인
            const mergedCells = worksheet['!merges'] || [];
            const range = worksheet['!ref'] || 'A1:Z100';
            
            console.log(`시트 범위: ${range}`);
            console.log(`병합된 셀 수: ${mergedCells.length}`);
            
            try {
                // 객체 형태로 변환 (헤더 포함)
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 'A',
                    raw: false,
                    defval: ''
                });
                
                // 원시 데이터 확인
                console.log(`시트 ${sheetName}의 행 수: ${jsonData.length}`);
                
                // 빈 시트 건너뛰기
                if (!jsonData || jsonData.length < 3) {
                    console.log(`건너뛴 시트: ${sheetName} (데이터 부족)`);
                    continue;
                }
                
                if (jsonData.length > 0) {
                    // 첫 몇 행의 데이터를 확인
                    for (let i = 0; i < Math.min(5, jsonData.length); i++) {
                        console.log(`행 ${i}:`, JSON.stringify(jsonData[i]));
                    }
                }
                
                // 이 시트의 데이터 파싱
                const sheetData = this.parseSheetData(jsonData, finalDate);
                
                // 유효한 데이터가 있는 경우만 추가
                if (sheetData && sheetData.length > 0) {
                    // 결과를 전체 데이터에 추가
                    allParsedData = [...allParsedData, ...sheetData];
                    
                    // 완료 메시지
                    console.log(`시트 ${sheetName} 파싱 완료: ${sheetData.length}개 항목 추가됨`);
                } else {
                    console.log(`시트 ${sheetName}: 유효한 데이터 없음`);
                }
            } catch (error) {
                console.error(`시트 ${sheetName} 처리 중 오류 발생:`, error);
                // 오류가 있더라도 다른 시트는 계속 처리
                continue;
            }
        }
        
        // 중복 데이터 제거 (동일 부재번호 + 날짜)
        const uniqueItems = new Map();
        for (const item of allParsedData) {
            const key = `${item.date}-${item.assemblyNumber}`;
            uniqueItems.set(key, item);
        }
        
        const finalData = Array.from(uniqueItems.values());
        console.log(`총 처리된 데이터 수: ${allParsedData.length}, 중복 제거 후: ${finalData.length}`);
        
        // 첫 10개 항목 확인
        if (finalData.length > 0) {
            console.log('처리된 데이터 샘플:');
            for (let i = 0; i < Math.min(10, finalData.length); i++) {
                console.log(`[${i+1}] ${finalData[i].date}, ${finalData[i].assemblyNumber}, ${finalData[i].quantity}`);
            }
        } else {
            console.warn('처리된 데이터가 없습니다!');
        }
        
        return finalData;
    }
    
    parseSheetData(jsonData, sheetDate) {
        const parsedData = [];
        
        // 헤더 행과 열 인덱스 찾기
        let assemblyColIndex = null;
        let productionColIndex = null;
        let headerRowIndex = -1;
        
        // 헤더 검색
        for (let i = 0; i < Math.min(20, jsonData.length); i++) {
            const row = jsonData[i];
            if (!row) continue;
            
            for (const key in row) {
                const value = String(row[key] || '');
                
                // 부재번호 열 찾기 (다양한 패턴 지원)
                if (value.includes('부재번호') || value.includes('품번') || 
                    value.includes('ASSY') || value.includes('자재코드')) {
                    assemblyColIndex = key;
                    headerRowIndex = i;
                    console.log(`부재번호 열 발견: ${key}, 값: ${value}`);
                }
                
                // 생산량 열 찾기 (다양한 패턴 지원)
                if (value.includes('생산량') || value.includes('생산수량') || 
                    value.includes('수량') || value.includes('투입수량')) {
                    productionColIndex = key;
                    console.log(`생산량 열 발견: ${key}, 값: ${value}`);
                }
            }
            
            // 헤더를 찾았으면 루프 종료
            if (assemblyColIndex && productionColIndex && headerRowIndex !== -1) {
                break;
            }
        }
        
        // 헤더를 찾지 못한 경우 데이터 구조 유추
        if (!assemblyColIndex || !productionColIndex) {
            console.log('자동 헤더 검색 실패, 데이터 구조 분석 시도...');
            
            // 모든 셀 데이터 확인 (처음 10행만)
            for (let i = 0; i < Math.min(10, jsonData.length); i++) {
                const row = jsonData[i];
                if (!row) continue;
                
                console.log(`Row ${i} data:`, JSON.stringify(row));
            }
            
            // 나라피씨 데이터 구조 특성상 특정 열에 부재번호가 있을 가능성 검사
            for (let i = 0; i < Math.min(20, jsonData.length); i++) {
                const row = jsonData[i];
                if (!row) continue;
                
                // 각 열을 확인하여 부재번호 패턴 찾기
                for (const key in row) {
                    const value = String(row[key] || '');
                    
                    // 부재번호 패턴 (XX-XXX-XXXX) 검사
                    if (/^\d{2}-\d{3}-\d{4}$/.test(value)) {
                        assemblyColIndex = key;
                        headerRowIndex = i - 1; // 헤더는 이 행 바로 위에 있을 가능성이 높음
                        console.log(`부재번호 패턴 발견 (${i}행, ${key}열): ${value}`);
                        break;
                    }
                }
                
                if (assemblyColIndex) break;
            }
            
            // 부재번호 열을 찾았으면 그 근처에서 수량 열 추정
            if (assemblyColIndex) {
                const colCode = assemblyColIndex.charCodeAt(0);
                // 부재번호 열로부터 오른쪽으로 2~5칸 사이에 수량 열이 있을 가능성이 높음
                for (let i = 2; i <= 5; i++) {
                    const possibleQuantityCol = String.fromCharCode(colCode + i);
                    productionColIndex = possibleQuantityCol;
                    console.log(`추정된 생산량 열: ${productionColIndex} (부재번호 열 ${assemblyColIndex}에서 ${i}칸)`);
                    break;
                }
            }
            
            // 여전히 찾지 못했다면 기본값 사용
            if (!assemblyColIndex) {
                assemblyColIndex = 'B';  // 두 번째 열 (B열)
                headerRowIndex = 0;      // 첫 번째 행을 헤더로 가정
                console.log('부재번호 열을 찾을 수 없어 기본값 사용: B열');
            }
            
            if (!productionColIndex) {
                productionColIndex = 'D';  // 네 번째 열 (D열)
                console.log('생산량 열을 찾을 수 없어 기본값 사용: D열');
            }
        }
        
        console.log(`헤더 행: ${headerRowIndex}, 부재번호 열: ${assemblyColIndex}, 생산량 열: ${productionColIndex}`);
        
        // 헤더 행도 찾지 못한 경우
        if (headerRowIndex === -1) {
            headerRowIndex = 0; // 첫 번째 행을 헤더로 가정
            console.log('헤더 행을 찾을 수 없어 기본값 사용: 0행');
        }
        
        // '사용금지' 같은 제외 시트 처리
        if (jsonData.length < 3 || (assemblyColIndex === 'B' && productionColIndex === 'D' && jsonData.length < 5)) {
            console.log('유효한 데이터가 없는 시트로 판단됩니다. 건너뜁니다.');
            return [];
        }
        
        // 데이터 행 처리 (헤더 다음 행부터)
        let processedCount = 0;
        let invalidCount = 0;
        let excludedCount = 0;
        let noQuantityCount = 0;
        
        // 열 헤더를 출력 (디버깅용)
        if (jsonData[headerRowIndex]) {
            const headerRow = jsonData[headerRowIndex];
            console.log('헤더 행 내용:', JSON.stringify(headerRow));
        }
        
        for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;
            
            const assemblyNumber = row[assemblyColIndex];
            let productionQuantity = row[productionColIndex];
            
            // 데이터 존재 여부 디버깅
            if (i < headerRowIndex + 10) {
                console.log(`행 ${i}, 부재번호: ${assemblyNumber}, 생산량: ${productionQuantity}`);
            }
            
            // 부재번호가 있고 제외 키워드가 아닌 경우만 처리
            if (assemblyNumber && !this.isExcludedRow(assemblyNumber)) {
                // 부재번호 형식 정리 (공백 제거, 특수문자 처리 등)
                const cleanAssemblyNumber = String(assemblyNumber).trim();
                
                // 생산량 처리 - 문자열 형태로 저장된 경우도 처리
                let quantity = 0;
                
                if (productionQuantity !== undefined && productionQuantity !== null) {
                    // 쉼표 제거 및 공백 제거
                    const cleanQuantity = String(productionQuantity).replace(/,/g, '').trim();
                    
                    // 숫자로 변환
                    quantity = parseFloat(cleanQuantity);
                    
                    // 변환 실패 또는 0인 경우
                    if (isNaN(quantity) || quantity <= 0) {
                        // 다른 열에서 수량을 찾아봄
                        for (const key in row) {
                            if (key !== assemblyColIndex && key !== productionColIndex) {
                                const value = row[key];
                                if (value && !isNaN(parseFloat(String(value).replace(/,/g, '')))) {
                                    const possibleQuantity = parseFloat(String(value).replace(/,/g, ''));
                                    if (possibleQuantity > 0) {
                                        quantity = possibleQuantity;
                                        console.log(`대체 수량 발견 (${i}행, ${key}열): ${quantity}`);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                
                // 최종 데이터 추가
                if (quantity > 0) {
                    parsedData.push({
                        date: sheetDate,
                        assemblyNumber: cleanAssemblyNumber,
                        quantity: quantity,
                        company: 'narapc',
                        completedDate: sheetDate,
                        // 표준화된 필드명 추가
                        AssemblyNumber: cleanAssemblyNumber,
                        Quantity: quantity,
                        CompletedDate: sheetDate,
                        Company: 'narapc'
                    });
                    processedCount++;
                    
                    // 처리된 첫 10개 항목 출력 (디버깅용)
                    if (processedCount <= 10) {
                        console.log(`처리된 데이터 ${processedCount}: ${cleanAssemblyNumber}, 수량: ${quantity}`);
                    }
                } else {
                    if (productionQuantity === undefined || productionQuantity === null) {
                        noQuantityCount++;
                    } else {
                        invalidCount++;
                    }
                }
            } else if (assemblyNumber && this.isExcludedRow(assemblyNumber)) {
                excludedCount++;
            }
        }
        
        console.log(`시트 처리 완료 - 처리된 레코드 수: ${processedCount}, 무효한 수량 레코드 수: ${invalidCount}, 수량 없음: ${noQuantityCount}, 제외된 레코드 수: ${excludedCount}`);
        
        return parsedData;
    }

    async processFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // 파일명에서 날짜 추출
                    const fileDate = this.extractDateFromFilename(file.name);
                    
                    // 모든 시트의 데이터 파싱
                    const parsedData = this.parseFactoryData(workbook, fileDate);
                    resolve(parsedData);
                } catch (error) {
                    console.error('파일 처리 중 오류:', error);
                    reject(error);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    async analyzeFiles() {
        if (selectedFiles.length === 0) {
            alert('분석할 파일을 선택해주세요.');
            return;
        }

        // 로딩 표시
        document.getElementById('loadingIndicator').classList.remove('d-none');
        document.getElementById('resultSection').classList.add('d-none');

        try {
            // 파일들을 병렬로 처리
            const filePromises = selectedFiles.map(file => this.processFile(file));
            const results = await Promise.all(filePromises);

            // 모든 데이터를 하나의 배열로 합치기
            allData = results.flat();

            // 날짜별로 정렬
            allData.sort((a, b) => a.date.localeCompare(b.date));

            // 날짜 필터 옵션 업데이트
            this.updateDateFilterOptions();

            // currentSort 변수 확인 및 초기화
            if (typeof window.currentSort === 'undefined') {
                window.currentSort = { column: 'date', direction: 'asc' };
                console.log('currentSort 변수가 초기화되었습니다.');
            }

            try {
                // 필터 적용 및 결과 표시
                applyFilters();
                this.displayFilteredData();
            } catch (filterError) {
                console.error('필터 적용 중 오류 발생:', filterError);
                // 필터 오류가 발생해도 기본 데이터 표시
                filteredData = [...allData];
                this.displayFilteredData();
            }

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

    updateDateFilterOptions() {
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

    displayFilteredData() {
        const tbody = document.querySelector('#resultTable tbody');
        tbody.innerHTML = '';
        
        // 소계 행을 제외한 데이터만 표시
        filteredData.forEach(item => {
            if (!item.assemblyNumber.toLowerCase().includes('소계') &&
                !item.assemblyNumber.toLowerCase().includes('합계') &&
                !item.assemblyNumber.toLowerCase().includes('total') &&
                !item.assemblyNumber.toLowerCase().includes('subtotal')) {
                this.addTableRow(item);
            }
        });
    }

    addTableRow(item) {
        const tbody = document.querySelector('#resultTable tbody');
        const row = document.createElement('tr');
        
        row.innerHTML = `
            <td>${item.CompletedDate || item.date}</td>
            <td>${item.AssemblyNumber || item.assemblyNumber}</td>
            <td>${item.Quantity || item.quantity}</td>
        `;
        
        tbody.appendChild(row);
    }
}

// 나라피씨 분석기 인스턴스 생성
const naraPCAnalyzer = new NaraPCAnalyzer();

// 나라피씨 분석 함수
function analyzeNaraPCFiles() {
    return naraPCAnalyzer.analyzeFiles();
}

// 나라피씨 데이터 파싱 함수 (common.js에서 호출)
window.parseNaraPCData = function(jsonData, fileDate) {
    return naraPCAnalyzer.parseFactoryData(jsonData, fileDate);
};