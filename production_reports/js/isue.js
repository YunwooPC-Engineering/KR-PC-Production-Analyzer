const XLSX = require('xlsx');
const path = require('path');

class IsueDataParser {
    constructor() {
        this.excludedKeywords = ['소계', '합계', 'total', 'subtotal', '제품명'];
        this.currentYear = new Date().getFullYear();
        this.startRow = 2; // 시작 행 (3번째 행부터)
    }

    // 날짜 추출 함수 수정
    extractDateFromFilename(filename) {
        try {
            // 1. 폴더 경로에서 연도 추출 (예: 25년 3월/0322)
            const yearPattern = /[\/\\](\d{2})[년][\/\\]/;
            const yearMatch = filename.match(yearPattern);
            let year = null;
            
            if (yearMatch) {
                // 2자리 연도를 4자리로 변환
                year = parseInt(`20${yearMatch[1]}`);
                
                // 미래 연도 확인 (현재 연도보다 2년 이상 미래는 오류로 판단)
                if (year > this.currentYear + 2) {
                    console.warn(`미래 연도 감지 (${year}), 현재 연도로 조정합니다.`);
                    year = this.currentYear;
                }
            }
            
            // 2. 파일명에서 날짜 형식 추출 (0322 등)
            let month = null;
            let day = null;
            
            // MMDD 형식 찾기
            const datePattern = /(\d{2})(\d{2})/;
            const dateMatch = filename.match(datePattern);
            
            if (dateMatch) {
                month = parseInt(dateMatch[1]);
                day = parseInt(dateMatch[2]);
                
                // 유효한 월/일인지 확인
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    // 현재 날짜 가져오기
                    const now = new Date();
                    const currentMonth = now.getMonth() + 1;
                    const currentDay = now.getDate();
                    
                    // 연도가 추출되지 않은 경우 현재 연도 사용
                    if (!year) {
                        year = this.currentYear;
                        
                        // 월/일이 현재보다 미래인 경우 작년으로 처리
                        if (month > currentMonth || (month === currentMonth && day > currentDay)) {
                            year = this.currentYear - 1;
                        }
                    }
                    
                    // 최종 날짜 문자열 생성 (YYYYMMDD 형식)
                    const formattedMonth = String(month).padStart(2, '0');
                    const formattedDay = String(day).padStart(2, '0');
                    
                    console.log(`날짜 추출 결과: ${year}${formattedMonth}${formattedDay} (경로: ${filename})`);
                    return `${year}${formattedMonth}${formattedDay}`;
                }
            }
            
            // 날짜 추출 실패 시 현재 날짜 반환
            const now = new Date();
            const currentMonth = String(now.getMonth() + 1).padStart(2, '0');
            const currentDay = String(now.getDate()).padStart(2, '0');
            console.warn(`날짜를 추출할 수 없어 현재 날짜 사용: ${this.currentYear}${currentMonth}${currentDay}`);
            return `${this.currentYear}${currentMonth}${currentDay}`;
        } catch (error) {
            console.error('날짜 추출 중 오류:', error);
            
            // 오류 발생 시 현재 날짜 반환
            const now = new Date();
            const year = now.getFullYear();
            const month = String(now.getMonth() + 1).padStart(2, '0');
            const day = String(now.getDate()).padStart(2, '0');
            return `${year}${month}${day}`;
        }
    }

    extractDateFromDocument(workbook) {
        try {
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            
            // 문서 내에서 날짜 찾기 시도
            // 다양한 셀을 확인해 날짜 형식 데이터 검색
            for (let row = 0; row < 10; row++) {
                for (let col = 0; col < 5; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                    const cell = firstSheet[cellAddress];
                    
                    if (cell && cell.v) {
                        const cellValue = cell.v.toString();
                        // 날짜 패턴 확인 (예: 2023-03-21, 2023/03/21, 23.03.21 등)
                        const datePattern = /(\d{2,4})[-./](\d{1,2})[-./](\d{1,2})/;
                        const match = cellValue.match(datePattern);
                        
                        if (match) {
                            let year = match[1];
                            if (year.length === 2) {
                                year = `20${year}`;
                            }
                            
                            const month = match[2].padStart(2, '0');
                            const day = match[3].padStart(2, '0');
                            
                            return `${year}${month}${day}`;
                        }
                    }
                }
            }
            
            return null;
        } catch (error) {
            console.error('문서에서 날짜 추출 중 오류:', error);
            return null;
        }
    }

    shouldExcludeRow(rowData) {
        if (!rowData || rowData.length === 0) return true;
        
        // 제외할 키워드가 포함된 행 확인
        for (const keyword of this.excludedKeywords) {
            if (rowData.some(cell => 
                cell && 
                typeof cell === 'string' && 
                cell.toLowerCase().includes(keyword.toLowerCase())
            )) {
                return true;
            }
        }
        
        return false;
    }

    findProductionQuantityColumn(rows) {
        // 헤더가 있는 행 (일반적으로 첫 번째 행)
        const headerRow = rows[0] || [];
        
        // 다양한 생산량 관련 헤더를 검색
        const possibleHeaders = ['생산잔량', '생산수량', '생산량', '수량'];
        
        // 각 가능한 헤더에 대해 검색
        for (const header of possibleHeaders) {
            for (let i = 0; i < headerRow.length; i++) {
                const cellValue = headerRow[i];
                if (cellValue && 
                    typeof cellValue === 'string' && 
                    cellValue.includes(header)) {
                    console.log(`생산량 열 인덱스 발견: ${i}, 헤더: ${cellValue}`);
                    return i;
                }
            }
        }
        
        // 특정 열 위치에 생산량이 있는지 확인 (경험적 접근)
        const potentialColumns = [5, 6, 11, 10, 8];
        for (const colIndex of potentialColumns) {
            if (colIndex < headerRow.length) {
                console.log(`생산량 열 추정: ${colIndex}, 헤더 값: ${headerRow[colIndex]}`);
                return colIndex;
            }
        }
        
        console.log('생산량 열을 찾을 수 없음. 기본값 6 사용');
        return 6; // 기본값
    }

    findAssemblyNumberColumn(rows) {
        // 헤더 행 확인 (여러 행에서 시도)
        for (let rowIndex = 0; rowIndex < Math.min(rows.length, 10); rowIndex++) {
            const headerRow = rows[rowIndex] || [];
            
            // '부재번호' 헤더 검색
            for (let i = 0; i < headerRow.length; i++) {
                const cellValue = headerRow[i];
                if (cellValue && 
                    typeof cellValue === 'string' && 
                    cellValue.includes('부재번호')) {
                    console.log(`부재번호 열 인덱스 발견: ${i}, 헤더: ${cellValue}`);
                    return i;
                }
            }
        }
        
        // 부재번호 패턴을 가진 셀이 있는지 확인 (여러 행에서)
        for (let rowIndex = 0; rowIndex < Math.min(rows.length, 30); rowIndex++) {
            const row = rows[rowIndex] || [];
            for (let i = 0; i < row.length; i++) {
                const cellValue = row[i];
                if (cellValue && 
                    typeof cellValue === 'string' && 
                    /^\d{2}-\d{3}-\d{4}$/.test(cellValue)) {
                    console.log(`부재번호 패턴을 가진 열 발견: ${i}, 값: ${cellValue}`);
                    return i;
                }
            }
        }
        
        console.log('부재번호 열을 찾을 수 없음. 기본값 2 사용');
        return 2; // 기본값
    }

    parseExcelData(workbook, filename) {
        try {
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            
            // 시트 데이터를 2D 배열로 변환
            const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            
            // 데이터 디버깅
            console.log('첫 번째 행 데이터:');
            if (rows.length > 0) console.log(rows[0]);
            console.log('두 번째 행 데이터:');
            if (rows.length > 1) console.log(rows[1]);
            console.log('세 번째 행 데이터:');
            if (rows.length > 2) console.log(rows[2]);
            
            // 40행까지 각 행의 셀 확인
            for (let i = 3; i < Math.min(rows.length, 40); i++) {
                const row = rows[i];
                if (row) {
                    // 각 주요 셀 확인
                    console.log(`행 ${i+1}, 열 1: "${row[0]}"`);
                    console.log(`행 ${i+1}, 열 2: "${row[1]}"`);
                    console.log(`행 ${i+1}, 열 3: "${row[2]}"`);
                    // 몇 개의 추가 열 확인
                    if (row.length > 5) console.log(`행 ${i+1}, 열 6: "${row[5]}"`);
                    if (row.length > 10) console.log(`행 ${i+1}, 열 11: "${row[10]}"`);
                    if (row.length > 14) console.log(`행 ${i+1}, 열 15: "${row[14]}"`);
                }
            }
            
            // 날짜 추출 (파일명에서 먼저 시도, 실패하면, 문서에서 시도)
            let date = this.extractDateFromFilename(filename);
            if (!date) {
                date = this.extractDateFromDocument(workbook);
                console.log('문서에서 추출한 날짜:', date);
            }
            
            // 날짜를 찾지 못한 경우 기본값 설정
            if (!date) {
                date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
                console.log('날짜를 찾을 수 없음. 기본값 사용:', date);
            }
            
            console.log('최종 사용 날짜:', date);
            
            // 생산량 열 인덱스 찾기
            const productionQuantityIndex = this.findProductionQuantityColumn(rows);
            console.log(`생산량 열 인덱스: ${productionQuantityIndex}`);
            
            // 부재번호 열 인덱스 찾기
            const assemblyNumberIndex = this.findAssemblyNumberColumn(rows);
            console.log(`부재번호 열 인덱스: ${assemblyNumberIndex}`);
            
            const parsedData = [];
            let processedCount = 0;
            let invalidCount = 0;
            let excludedCount = 0;
            
            // 데이터 행 처리 (시작 행부터)
            for (let i = this.startRow; i < rows.length; i++) {
                const row = rows[i];
                
                // 빈 행 스킵
                if (!row || row.length === 0) continue;
                
                // 제외 키워드가 있는 행 스킵
                if (this.shouldExcludeRow(row)) {
                    excludedCount++;
                    continue;
                }
                
                // 부재번호와 생산량 추출
                const assemblyNumber = row[assemblyNumberIndex];
                const quantity = row[productionQuantityIndex];
                
                // 유효한 데이터만 추가
                if (assemblyNumber && quantity && !isNaN(Number(quantity)) && Number(quantity) > 0) {
                    // 날짜를 포맷팅하여 필드에 저장
                    const formattedDate = date ? date : '';
                    
                    parsedData.push({
                        date: formattedDate,               // 원래 형식의 날짜 (YYYYMMDD)
                        assemblyNumber: assemblyNumber.toString(),  // 부재번호
                        quantity: Number(quantity),        // 생산량
                        company: 'jinsungpc',              // 회사 식별자
                        completedDate: formattedDate,      // 완료일(CompletedDate) - DB 호환성 
                        AssemblyNumber: assemblyNumber.toString(),  // 대문자 버전 (표준화)
                        Quantity: Number(quantity),        // 대문자 버전 (표준화)
                        CompletedDate: formattedDate       // 대문자 버전 (표준화)
                    });
                    processedCount++;
                } else {
                    invalidCount++;
                }
            }
            
            console.log(`처리된 레코드 수: ${processedCount}`);
            console.log(`무효한 레코드 수: ${invalidCount}`);
            console.log(`제외된 레코드 수: ${excludedCount}`);
            
            // 파싱된 데이터의 일부 샘플 표시
            console.log('파싱된 데이터 구조 (첫 10행):');
            console.log(JSON.stringify(parsedData.slice(0, 10), null, 2));
            
            return parsedData;
        } catch (error) {
            console.error('Excel 데이터 파싱 중 오류:', error);
            return [];
        }
    }
}

// 이수이앤씨 파서 함수
function parseIsueData(filePath) {
    try {
        const filename = path.basename(filePath);
        const workbook = XLSX.readFile(filePath);
        
        const parser = new IsueDataParser();
        const parsedData = parser.parseExcelData(workbook, filename);
        
        console.log(`이수이앤씨 데이터 파싱 완료. ${parsedData.length}개 레코드 처리됨.`);
        return parsedData;
    } catch (error) {
        console.error('이수이앤씨 데이터 파싱 중 오류:', error);
        return [];
    }
}

module.exports = {
    parseIsueData
};