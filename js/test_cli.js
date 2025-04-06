// test_cli.js 파일 - 커맨드 라인에서 엑셀 파일을 테스트하기 위한 스크립트
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// 파일 경로 가져오기
const filePath = process.argv[2];

if (!filePath) {
    console.error('파일 경로를 지정해주세요.');
    console.error('사용법: node js/test_cli.js "파일경로.xlsx"');
    process.exit(1);
}

console.log(`파일 처리중: ${filePath}`);

// 파일명 및 회사 이름 추출
const fileName = path.basename(filePath);
let company = 'unknown';

if (fileName.toLowerCase().includes('진성')) {
    company = 'jinsungpc';
} else if (fileName.toLowerCase().includes('이수')) {
    company = 'isue';
} else if (fileName.toLowerCase().includes('지산')) {
    company = 'jisan';
} else if (fileName.toLowerCase().includes('나라')) {
    company = 'narapc';
} else {
    company = 'jinsungpc'; // 기본값
}

// 파일 이름에서 날짜 추출
function extractDateFromFilename(filename) {
    // 현재 날짜 구하기
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1;
    const currentDay = now.getDate();
    
    // 최종 사용할 날짜 값
    let finalYear = currentYear;
    let finalMonth = null;
    let finalDay = null;
    
    // 1. 경로에서 연도 패턴 확인 ("25년 3월" 같은 패턴)
    const yearFromPathPattern = /[\/\\](\d{2})[년][\/\\]/;
    const yearMatch = filename.match(yearFromPathPattern);
    
    if (yearMatch) {
        // 2자리 연도를 감지한 경우 현재 연도 사용 (미래 날짜 문제 방지)
        console.log(`파일 경로에서 연도 패턴 감지: ${yearMatch[1]}년, 현재 연도(${currentYear})를 사용합니다.`);
        finalYear = currentYear;
        
        // 경로에서 월 추출 시도
        const monthPattern = /[\/\\]\d{2}년\s*(\d{1,2})월[\/\\]/;
        const monthMatch = filename.match(monthPattern);
        if (monthMatch) {
            const monthValue = parseInt(monthMatch[1]);
            if (monthValue >= 1 && monthValue <= 12) {
                finalMonth = monthValue;
            }
        }
    }
    
    // 2. 파일명에서 MMDD 패턴 추출
    const shortDatePattern = /(\d{2})(\d{2})/;
    const dateMatch = filename.match(shortDatePattern);
    
    if (dateMatch) {
        const month = parseInt(dateMatch[1]);
        const day = parseInt(dateMatch[2]);
        
        // 유효한 월/일인지 확인
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            finalMonth = month;
            finalDay = day;
            
            // 월/일이 현재보다 미래인 경우 작년으로 설정
            if (finalYear === currentYear && 
                (month > currentMonth || (month === currentMonth && day > currentDay))) {
                console.warn(`현재 날짜 이후의 날짜 감지, 연도를 작년으로 조정: ${finalYear-1}`);
                finalYear = currentYear - 1;
            }
        }
    }
    
    // 3. 파일명에서 YYYYMMDD 패턴 추출
    if (!finalDay) {
        const fullDatePattern = /(\d{4})(\d{2})(\d{2})/;
        let match = filename.match(fullDatePattern);
        
        if (match) {
            const year = parseInt(match[1]);
            const month = parseInt(match[2]);
            const day = parseInt(match[3]);
            
            // 유효한 날짜 범위인지 확인
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                // 2년 이상 미래 연도는 현재 연도로 조정
                if (yearMatch) {
                    // 경로에서 연도가 감지된 경우 현재 연도 우선 적용
                    console.warn(`경로에서 연도 감지됨. 파일명의 연도(${year}) 대신 현재 연도(${currentYear}) 사용`);
                    finalYear = currentYear;
                } else if (year > currentYear + 2) {
                    console.warn(`미래 연도 감지: ${year}, 현재 연도로 조정합니다.`);
                    finalYear = currentYear;
                } else {
                    finalYear = year;
                }
                
                finalMonth = month;
                finalDay = day;
            }
        }
    }
    
    // 최종 날짜 생성
    if (finalMonth && finalDay) {
        const formattedMonth = String(finalMonth).padStart(2, '0');
        const formattedDay = String(finalDay).padStart(2, '0');
        const dateStr = `${finalYear}${formattedMonth}${formattedDay}`;
        console.log(`최종 추출된 날짜: ${dateStr}`);
        return dateStr;
    }
    
    // 날짜를 찾지 못한 경우 현재 날짜 사용
    const month = String(currentMonth).padStart(2, '0');
    const day = String(currentDay).padStart(2, '0');
    console.warn(`날짜를 찾을 수 없어 현재 날짜를 사용합니다: ${finalYear}${month}${day}`);
    return `${finalYear}${month}${day}`;
}

// 문서에서 날짜 추출
function extractDateFromDocument(data) {
    // 첫 몇 행에서 날짜 형식 찾기
    for (let i = 0; i < Math.min(5, data.length); i++) {
        const row = data[i];
        if (!row || !Array.isArray(row)) continue;
        
        // 각 셀을 검사
        for (let j = 0; j < row.length; j++) {
            const value = row[j];
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

// 부재번호 패턴 검사 (xx-xxx-xxxx)
function isPotentialAssemblyNumber(value) {
    if (!value) return false;
    const strValue = String(value);
    return /\d{2}-\d{3}-\d{4}/.test(strValue);
}

// 파일 읽기
try {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // 헤더포함 모든 데이터 JSON 변환
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // 파일명에서 날짜 추출
    const fileDate = extractDateFromFilename(fileName);

    // 처음 몇 행 출력
    console.log('첫 번째 행 데이터:');
    console.log(jsonData[0]);
    console.log('두 번째 행 데이터:');
    console.log(jsonData[1]);
    console.log('세 번째 행 데이터:');
    console.log(jsonData[2]);

    // 특정 셀 내용 확인 (행 4-40, 열 A, B, C, F, K, O)
    for (let row = 3; row < Math.min(40, jsonData.length); row++) {
        const currentRow = jsonData[row];
        if (!currentRow) continue;
        
        console.log(`행 ${row+1}, 열 1: "${currentRow[0]}"`);
        console.log(`행 ${row+1}, 열 2: "${currentRow[1]}"`);
        console.log(`행 ${row+1}, 열 3: "${currentRow[2]}"`);
        if (currentRow[5] !== undefined) console.log(`행 ${row+1}, 열 6: "${currentRow[5]}"`);
        if (currentRow[10] !== undefined) console.log(`행 ${row+1}, 열 11: "${currentRow[10]}"`);
        if (currentRow[14] !== undefined) console.log(`행 ${row+1}, 열 15: "${currentRow[14]}"`);
    }

    // 회사별 파싱
    let parsedData = [];
    
    if (company === 'jisan') {
        parsedData = parseJisanData(jsonData, fileDate);
    } else if (company === 'isue') {
        parsedData = parseIsueDataLocal(jsonData, fileDate);
    } else {
        // 기본 파서는 진성피씨로 사용
        parsedData = parseJinsungPCData(jsonData, fileDate);
    }

    // 결과 출력
    console.log('파싱된 데이터 구조 (첫 10행):');
    console.log(JSON.stringify(parsedData.slice(0, 10), null, 2));

    console.log(`${company} 데이터 파싱 완료. ${parsedData.length}개 레코드 처리됨.`);

    // 결과를 파일로 저장
    const outputDir = path.dirname(filePath);
    const outputFileName = `parsed_진성 생산 실적_${fileDate}.json`;
    const outputPath = path.join(outputDir, outputFileName);

    fs.writeFileSync(outputPath, JSON.stringify(parsedData, null, 2));
    console.log(`결과 저장됨: ${outputPath}`);

    // 모든 파일 결과를 저장
    const allResultsPath = path.join(outputDir, 'all_results.json');
    let allResults = {};

    // 기존 결과 파일이 있으면 로드
    if (fs.existsSync(allResultsPath)) {
        try {
            allResults = JSON.parse(fs.readFileSync(allResultsPath, 'utf8'));
        } catch (error) {
            console.error('기존 결과 파일을 로드하는 중 오류 발생:', error);
            allResults = {};
        }
    }

    // 결과 추가
    if (!allResults.files) allResults.files = [];
    if (!allResults.results) allResults.results = [];
    if (!allResults.dateCount) allResults.dateCount = {};

    // 파일 정보 추가
    allResults.files.push({
        file: fileName,
        recordCount: parsedData.length,
        date: fileDate,
        company: company
    });

    // 데이터 추가
    allResults.results = allResults.results.concat(parsedData);

    // 날짜별 카운트 업데이트
    if (!allResults.dateCount[fileDate]) {
        allResults.dateCount[fileDate] = parsedData.length;
    } else {
        allResults.dateCount[fileDate] += parsedData.length;
    }

    // 저장
    fs.writeFileSync(allResultsPath, JSON.stringify(allResults, null, 2));

    // 처리된 레코드 수 출력
    console.log(`처리된 레코드 수: ${parsedData.length}`);

    // 데이터 샘플 출력 (처음 5개)
    console.log('데이터 샘플 (처음 5개):');
    console.log(parsedData.slice(0, 5));

    // 데이터 샘플 출력 (마지막 5개)
    console.log('데이터 샘플 (마지막 5개):');
    console.log(parsedData.slice(-5));

    // 처리 결과 요약
    console.log('\n=== 처리 결과 요약 ===');
    console.log(`처리된 총 파일 수: ${allResults.files.length}`);
    console.log(`실패한 파일 수: ${allResults.files.filter(f => f.recordCount === 0).length}`);

    // 날짜별 레코드 수 출력
    console.log('날짜별 레코드 수:');
    for (const date in allResults.dateCount) {
        console.log(`- ${date}: ${allResults.dateCount[date]}개`);
    }

    console.log(`총 레코드 수: ${allResults.results.length}`);
    console.log(`모든 결과 저장 경로: ${allResultsPath}`);

} catch (error) {
    console.error('파일 처리 중 오류 발생:', error);
    process.exit(1);
}

// 지산개발 데이터 파싱 함수
function parseJisanData(jsonData, fileDate) {
    // 문서에서 날짜 추출
    const documentDate = extractDateFromDocument(jsonData);
    
    // 최종 사용할 날짜 결정
    const finalDate = documentDate || fileDate;
    console.log('최종 사용 날짜:', finalDate);
    
    const parsedData = [];
    let headerRowIndex = -1;
    let assemblyColIndex = -1;
    let productionColIndex = -1;
    const excludedKeywords = ['소계', '합계', 'total', 'subtotal'];
    
    // 헤더 검색
    for (let i = 0; i < Math.min(20, jsonData.length); i++) {
        const row = jsonData[i];
        if (!row) continue;
        
        // 헤더 행 찾기
        for (let j = 0; j < row.length; j++) {
            const value = String(row[j] || '');
            
            // 제품번호로 열 식별
            if (value.includes('제품번호')) {
                assemblyColIndex = j;
                headerRowIndex = i;
            }
            
            // 생산량 관련 열 찾기 (생산량, 수량, 금일)
            if (value.includes('생산량') || 
                (value.includes('생산') && value.includes('금일'))) {
                productionColIndex = j;
            }
        }
        
        // 헤더를 찾았으면 루프 종료
        if (assemblyColIndex !== -1 && headerRowIndex !== -1) {
            break;
        }
    }
    
    // 생산량 열을 찾지 못한 경우 더 자세히 검색
    if (productionColIndex === -1) {
        // 생산량 > 수량 > 금일 순서로 검색
        for (let i = 0; i < Math.min(20, jsonData.length); i++) {
            const row = jsonData[i];
            if (!row) continue;
            
            for (let j = 0; j < row.length; j++) {
                const value = String(row[j] || '');
                
                if (value.includes('생산량')) {
                    for (let k = j; k < row.length; k++) {
                        const subValue = String(row[k] || '');
                        if (subValue.includes('수량') || subValue.includes('금일')) {
                            productionColIndex = k;
                            break;
                        }
                    }
                    
                    if (productionColIndex !== -1) break;
                } else if (value.includes('설계량')) {
                    // 설계량 다음에 생산량 열이 있을 수 있음
                    for (let k = 1; k < 10; k++) {
                        if (j+k < row.length && row[j+k]) {
                            const subValue = String(row[j+k] || '');
                            if (subValue.includes('생산량') || 
                                subValue.includes('수량') || 
                                subValue.includes('금일')) {
                                productionColIndex = j+k;
                                break;
                            }
                        }
                    }
                }
            }
            
            if (productionColIndex !== -1) break;
        }
    }
    
    // 여전히 찾지 못한 경우, 설계량 열 검색 후 6열 뒤를 기본값으로 설정
    if (productionColIndex === -1) {
        for (let i = 0; i < Math.min(20, jsonData.length); i++) {
            const row = jsonData[i];
            if (!row) continue;
            
            for (let j = 0; j < row.length; j++) {
                if (row[j] && String(row[j]).includes('설계량')) {
                    productionColIndex = j + 6; // 일반적으로 설계량에서 6열 뒤에 생산량이 있다고 가정
                    console.log('생산량 열을 찾을 수 없음. 기본값', productionColIndex, '사용');
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
        if (assemblyNumber && isPotentialAssemblyNumber(assemblyNumber)) {
            console.log('제품번호 패턴을 가진 열 발견:', i, '값:', assemblyNumber);
        }
        
        // 제품번호와 생산량이 있고, 제외 키워드가 없는 행만 처리
        if (assemblyNumber && productionQuantity !== undefined) {
            // 소계, 합계 등의 행 제외
            if (assemblyNumber && excludedKeywords.some(keyword => 
                String(assemblyNumber).toLowerCase().includes(keyword))) {
                excludedCount++;
                continue;
            }
            
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
        }
    }
    
    console.log('처리된 레코드 수:', processedCount);
    console.log('무효한 레코드 수:', invalidCount);
    console.log('제외된 레코드 수:', excludedCount);
    
    return parsedData;
}

// 이수이앤씨 데이터 파싱 함수
function parseIsueDataLocal(jsonData, fileDate) {
    // 문서에서 날짜 추출
    const documentDate = extractDateFromDocument(jsonData);
    
    // 최종 사용할 날짜 결정
    const finalDate = documentDate || fileDate;
    console.log('최종 사용 날짜:', finalDate);
    
    const parsedData = [];
    let headerRowIndex = -1;
    let assemblyColIndex = -1;
    let quantityColIndex = -1;
    const excludedKeywords = ['소계', '합계', 'total', 'subtotal', '제품명'];
    
    // 헤더 행 찾기
    for (let i = 0; i < Math.min(20, jsonData.length); i++) {
        const row = jsonData[i];
        if (!row) continue;
        
        for (let j = 0; j < row.length; j++) {
            const cellValue = String(row[j] || '');
            
            // 부재번호 열 식별
            if (cellValue.includes('부재번호') || cellValue.includes('품번') || 
                cellValue.includes('제품번호') || cellValue.includes('자재번호')) {
                assemblyColIndex = j;
                headerRowIndex = i;
            }
            
            // 생산량/수량 열 식별
            if (cellValue.includes('생산량') || cellValue.includes('수량') || 
                cellValue.includes('생산수량') || cellValue.includes('생산잔량')) {
                quantityColIndex = j;
            }
        }
        
        if (headerRowIndex !== -1 && assemblyColIndex !== -1) break;
    }
    
    // 수량 열을 찾지 못한 경우, 더 넓은 범위에서 검색
    if (quantityColIndex === -1 && headerRowIndex !== -1) {
        const row = jsonData[headerRowIndex];
        for (let j = 0; j < row.length; j++) {
            const cellValue = String(row[j] || '').toLowerCase();
            if (cellValue.includes('생산') || cellValue.includes('수량')) {
                quantityColIndex = j;
                break;
            }
        }
    }
    
    // 여전히 찾지 못한 경우, 마지막 시도로 특정 위치 검사
    if (quantityColIndex === -1 && assemblyColIndex !== -1) {
        quantityColIndex = assemblyColIndex + 5; // 일반적으로 부재번호로부터 5칸 떨어진 곳에 생산량이 있다고 가정
        console.log('생산량 열을 찾을 수 없음. 기본값 설정:', quantityColIndex);
    }
    
    console.log('생산량 열 인덱스:', quantityColIndex);
    console.log('부재번호 열 인덱스:', assemblyColIndex);
    
    // 헤더 행을 찾지 못한 경우
    if (headerRowIndex === -1 || assemblyColIndex === -1 || quantityColIndex === -1) {
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
        const productionQuantity = row[quantityColIndex];
        
        // 부재번호가 유효한지 검사
        if (assemblyNumber && isPotentialAssemblyNumber(assemblyNumber)) {
            console.log('부재번호 패턴을 가진 열 발견:', i, '값:', assemblyNumber);
        }
        
        // 소계, 합계 등의 행 제외
        if (assemblyNumber && excludedKeywords.some(keyword => 
            String(assemblyNumber).toLowerCase().includes(keyword))) {
            excludedCount++;
            continue;
        }
        
        // 부재번호와 생산량이 있는 행만 처리
        if (assemblyNumber && productionQuantity !== undefined && 
            productionQuantity !== null && productionQuantity !== '') {
            
            // 생산량이 숫자인지 확인
            const quantity = parseFloat(productionQuantity);
            if (!isNaN(quantity) && quantity > 0) {
                parsedData.push({
                    date: finalDate,
                    assemblyNumber: String(assemblyNumber).trim(),
                    quantity: quantity,
                    company: 'isue',
                    completedDate: finalDate,
                    AssemblyNumber: String(assemblyNumber).trim(),
                    Quantity: quantity,
                    CompletedDate: finalDate
                });
                processedCount++;
            } else {
                invalidCount++;
            }
        }
    }
    
    console.log('처리된 레코드 수:', processedCount);
    console.log('무효한 레코드 수:', invalidCount);
    console.log('제외된 레코드 수:', excludedCount);
    
    return parsedData;
}

// 진성피씨 데이터 파싱 함수
function parseJinsungPCData(jsonData, fileDate) {
    const parsedData = [];
    let colHeaders = [];
    let headerRowIndex = -1;
    let assemblyColIndex = -1;
    let quantityColIndex = -1;

    // 헤더 행 찾기
    for (let i = 0; i < Math.min(20, jsonData.length); i++) {
        const row = jsonData[i];
        if (!row) continue;

        // 부재번호, 수량 관련 헤더 검색
        for (let j = 0; j < row.length; j++) {
            const cell = String(row[j] || '').toLowerCase();
            if (cell.includes('부재번호') || cell.includes('부재 번호') || cell.includes('품번') || 
                cell.includes('제품번호') || cell.includes('자재번호')) {
                assemblyColIndex = j;
                headerRowIndex = i;
                colHeaders = row;
                break;
            }
        }

        if (headerRowIndex !== -1) break;
    }

    // 헤더를 찾지 못한 경우
    if (headerRowIndex === -1) {
        console.error('헤더 행을 찾을 수 없습니다.');
        return parsedData;
    }

    // 수량 열 찾기
    for (let j = 0; j < colHeaders.length; j++) {
        const header = String(colHeaders[j] || '').toLowerCase();
        if (header.includes('수량') || header.includes('생산량') || header.includes('생산수량') || 
            header.includes('생산 수량') || header.includes('계획수량')) {
            quantityColIndex = j;
            break;
        }
    }

    // 수량 열을 찾지 못한 경우, 부재번호 열 옆에 있다고 가정
    if (quantityColIndex === -1 && assemblyColIndex !== -1) {
        quantityColIndex = assemblyColIndex + 1;
    }

    // 데이터 행 처리 (헤더 다음 행부터)
    let processedCount = 0;
    let invalidCount = 0;
    let excludedCount = 0;

    for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row) continue;

        const assemblyNumber = row[assemblyColIndex];
        const quantity = row[quantityColIndex];

        // 부재번호와 수량이 있는 행만 처리
        if (assemblyNumber && quantity !== undefined && quantity !== null && quantity !== '') {
            // 소계, 합계 등의 행 제외
            const assemblyStr = String(assemblyNumber).toLowerCase();
            if (assemblyStr.includes('소계') || assemblyStr.includes('합계') || 
                assemblyStr.includes('total') || assemblyStr.includes('계') || 
                assemblyStr.includes('subtotal')) {
                excludedCount++;
                continue;
            }

            // 수량이 숫자인지 확인
            const parsedQuantity = parseFloat(quantity);
            if (!isNaN(parsedQuantity) && parsedQuantity > 0) {
                parsedData.push({
                    date: fileDate,
                    assemblyNumber: String(assemblyNumber).trim(),
                    quantity: parsedQuantity,
                    company: 'jinsungpc',
                    completedDate: fileDate,
                    AssemblyNumber: String(assemblyNumber).trim(),
                    Quantity: parsedQuantity,
                    CompletedDate: fileDate
                });
                processedCount++;
            } else {
                invalidCount++;
            }
        }
    }

    console.log(`처리된 레코드 수: ${processedCount}`);
    console.log(`무효한 레코드 수: ${invalidCount}`);
    console.log(`제외된 레코드 수: ${excludedCount}`);

    return parsedData;
}