// 지산개발 생산일보 분석
class JisanAnalyzer {
    constructor() {
        this.factoryName = '지산개발';
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

    // 문서에서 날짜 추출
    extractDateFromDocument(data) {
        // 첫 몇 행에서 날짜 형식 찾기
        for (let i = 0; i < Math.min(5, data.length); i++) {
            const row = data[i];
            if (!row) continue;
            
            // 각 셀을 검사
            for (const key in row) {
                const value = row[key];
                if (!value) continue;
                
                // YYYY-MM-DD 형식 검사
                const datePattern = /\d{4}-\d{2}-\d{2}/;
                const match = String(value).match(datePattern);
                if (match) {
                    // 하이픈 제거
                    return match[0].replace(/-/g, '');
                }
            }
        }
        return null;
    }

    // 제외할 키워드 체크
    isExcludedRow(assemblyNumber) {
        if (!assemblyNumber) return true;
        
        const lowerStr = String(assemblyNumber).toLowerCase();
        return this.excludedKeywords.some(keyword => lowerStr.includes(keyword));
    }

    parseFactoryData(jsonData, fileDate) {
        const parsedData = [];
        
        // 문서에서 날짜 찾기
        const documentDate = this.extractDateFromDocument(jsonData);
        
        // 최종 사용할 날짜 결정
        const finalDate = documentDate || fileDate;
        console.log('최종 사용 날짜:', finalDate);
        
        // 헤더 행과 열 인덱스 찾기
        let assemblyColIndex = -1;
        let productionColIndex = -1;
        let headerRowIndex = -1;
        let cumulativeColIndex = -1;  // 누계 열 인덱스
        
        // 헤더 검색
        for (let i = 0; i < Math.min(20, jsonData.length); i++) {
            const row = jsonData[i];
            if (!row) continue;
            
            // 헤더 행 찾기
            for (const key in row) {
                const value = String(row[key] || '');
                
                // 제품번호로 열 식별
                if (value.includes('제품번호')) {
                    assemblyColIndex = key;
                    headerRowIndex = i;
                }
                
                // 생산량 열 찾기
                if (value.includes('생산량')) {
                    productionColIndex = key;
                }
                
                // 누계 열 찾기
                if (value.includes('누계')) {
                    cumulativeColIndex = key;
                }
            }
            
            // 헤더를 찾았으면 루프 종료
            if (assemblyColIndex !== -1 && headerRowIndex !== -1) {
                break;
            }
        }
        
        // 생산량 열을 찾지 못한 경우 먼저 누계 열을 기준으로 찾기
        if (productionColIndex === -1 && cumulativeColIndex !== -1) {
            // 누계 열이 있다면 그 왼쪽에 생산량이 있을 가능성이 높음
            const cumulativeCharCode = cumulativeColIndex.charCodeAt(0);
            productionColIndex = String.fromCharCode(cumulativeCharCode - 4);  // 일반적으로 누계 4칸 전이 생산량
            console.log('누계 열 기준으로 생산량 열 추정:', productionColIndex);
        }
        
        // 여전히 생산량 열을 찾지 못한 경우 더 넓은 검색
        if (productionColIndex === -1) {
            // 설계량 열을 찾고 그 오른쪽에서 생산량 또는 수량을 검색
            for (let i = 0; i < Math.min(20, jsonData.length); i++) {
                const row = jsonData[i];
                if (!row) continue;
                
                for (const key in row) {
                    const value = String(row[key] || '');
                    
                    if (value.includes('설계량')) {
                        // 설계량 다음에 생산량 열을 찾기 위해 여러 칸 검색
                        for (let j = 1; j <= 10; j++) {
                            const nextColKey = String.fromCharCode(key.charCodeAt(0) + j);
                            if (row[nextColKey]) {
                                const subValue = String(row[nextColKey] || '');
                                if (subValue.includes('생산량') || 
                                    subValue.includes('수량') || 
                                    subValue.includes('금일')) {
                                    productionColIndex = nextColKey;
                                    break;
                                }
                            }
                        }
                        
                        if (productionColIndex === -1) {
                            // 생산량을 찾지 못했다면 설계량에서 6칸 떨어진 곳을 기본값으로 설정
                            const keyCode = key.charCodeAt(0);
                            productionColIndex = String.fromCharCode(keyCode + 6);
                            console.log('생산량 열을 찾을 수 없음. 기본값', productionColIndex, '사용');
                        }
                        
                        break;
                    }
                }
                
                if (productionColIndex !== -1) break;
            }
        }
        
        console.log('생산량 열 인덱스:', productionColIndex);
        console.log('제품번호 열 인덱스:', assemblyColIndex);
        
        // 헤더 행을 찾지 못한 경우
        if (headerRowIndex === -1 || assemblyColIndex === -1 || productionColIndex === -1) {
            console.error('헤더 행 또는 필수 열을 찾을 수 없습니다.');
            return parsedData;
        }
        
        // 데이터 행 처리 (헤더 다음 행부터)
        let processedCount = 0;
        let invalidCount = 0;
        let excludedCount = 0;
        
        for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;
            
            const assemblyNumber = row[assemblyColIndex];
            const productionQuantity = row[productionColIndex];
            
            // 제품번호가 유효한지 검사
            if (assemblyNumber && this.isPotentialAssemblyNumber(assemblyNumber)) {
                console.log('제품번호 패턴을 가진 열 발견:', i, '값:', assemblyNumber);
            }
            
            // 제품번호와 생산량이 있고, 제외 키워드가 없는 행만 처리
            if (assemblyNumber && productionQuantity && !this.isExcludedRow(assemblyNumber)) {
                // 생산량이 숫자인지 확인
                const quantity = parseFloat(productionQuantity);
                if (!isNaN(quantity) && quantity > 0) {
                    parsedData.push({
                        date: finalDate,
                        assemblyNumber: String(assemblyNumber).trim(),
                        quantity: quantity,
                        company: 'jisan',
                        completedDate: finalDate,
                        AssemblyNumber: String(assemblyNumber).trim(),
                        Quantity: quantity,
                        CompletedDate: finalDate
                    });
                    processedCount++;
                } else {
                    invalidCount++;
                }
            } else if (assemblyNumber && this.isExcludedRow(assemblyNumber)) {
                excludedCount++;
            }
        }
        
        console.log('처리된 레코드 수:', processedCount);
        console.log('무효한 레코드 수:', invalidCount);
        console.log('제외된 레코드 수:', excludedCount);
        
        return parsedData;
    }
    
    // 부재번호 패턴 검사 (xx-xxx-xxxx)
    isPotentialAssemblyNumber(value) {
        if (!value) return false;
        const strValue = String(value);
        return /\d{2}-\d{3}-\d{4}/.test(strValue);
    }

    async processFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 'A' });

                    // 파일명에서 날짜 추출
                    const fileDate = this.extractDateFromFilename(file.name);

                    // 데이터 파싱
                    const parsedData = this.parseFactoryData(jsonData, fileDate);
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
    
    // FactoryAnalyzer 클래스에서 상속받던 메소드들 추가
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

            // 필터 적용 및 결과 표시
            applyFilters();
            this.displayFilteredData();

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
            <td>${item.date}</td>
            <td>${item.assemblyNumber}</td>
            <td>${item.quantity}</td>
        `;
        
        tbody.appendChild(row);
    }
}

// 지산개발 분석기 인스턴스 생성
const jisanAnalyzer = new JisanAnalyzer();

// 지산개발 분석 함수
function analyzeJisanFiles() {
    return jisanAnalyzer.analyzeFiles();
}

// 지산개발 데이터 파싱 함수 (common.js에서 호출)
window.parseJisanData = function(jsonData, fileDate) {
    return jisanAnalyzer.parseFactoryData(jsonData, fileDate);
}; 