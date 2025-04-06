// 진성피씨 업체 분석용 스크립트

/**
 * 진성피씨 생산일보 파싱 함수
 */

// 진성피씨 데이터 파싱 함수
function parseJinsungPCData(jsonData, fileDate) {
    // 파일명에서 추출한 날짜 기준값 (형식: YYYYMMDD)
    // 날짜 강제 변환 로직 제거
    let extractedDate = String(fileDate || '');
    
    console.log('진성피씨 데이터 파싱 시작, 파일 날짜:', extractedDate);
    console.log('데이터 행 수:', jsonData.length);
    
    // 엑셀 내용에서 날짜 찾기 (더 정확한 날짜 추출)
    let documentDate = findDateInDocument(jsonData);
    console.log('문서에서 찾은 날짜:', documentDate);
    
    // 최종 사용할 날짜 결정 (우선순위: 문서내 날짜 > 파일명 날짜)
    let finalDate = documentDate || extractedDate;
    console.log('최종 사용할 날짜:', finalDate);
    
    // 날짜가 없는 경우 경고
    if (!finalDate) {
        console.warn('⚠️ 날짜를 찾을 수 없습니다! 현재 날짜를 사용합니다.');
        // 현재 날짜를 YYYYMMDD 형식으로 사용
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        finalDate = `${year}${month}${day}`;
    }
    
    // 샘플 데이터 출력 (처음 10행)
    console.log('\n=== 데이터 구조 샘플 (처음 10행) ===');
    for (let i = 0; i < Math.min(10, jsonData.length); i++) {
        console.log(`행 ${i+1}:`, jsonData[i]);
    }
    
    // 결과 데이터 배열
    const results = [];
    
    // 엑셀 구조 분석 (진성피씨 기준)
    // 2행에 헤더가 있고, 4행부터 데이터 시작
    let headerRowIndex = 1; // 2번째 행이 헤더 (0-based index)
    let dataStartIndex = 3; // 4번째 행부터 데이터 시작
    
    // 열 인덱스 (진성피씨 엑셀 파일 기준 하드코딩)
    let assemblyColIndex = 1; // 부재명 열 (2번째 열)
    let designQuantityIndex = 2; // 설계수량 열 (3번째 열)
    let productionQuantityIndex = 5; // 생산잔량(EA) 열 (6번째 열)
    
    // 열 인덱스를 확인하기 위해 헤더 행 데이터 출력
    if (jsonData.length > headerRowIndex) {
        const headerRow = jsonData[headerRowIndex];
        console.log('헤더 행 내용:', headerRow);
        
        if (headerRow && Array.isArray(headerRow)) {
            // 열 이름 출력
            for (let i = 0; i < headerRow.length; i++) {
                if (headerRow[i]) {
                    console.log(`열 ${i+1}: "${headerRow[i]}"`);
                    
                    // 부재명과 생산수량 열 확인 (case insensitive)
                    const cellStr = String(headerRow[i]).toLowerCase();
                    if (cellStr.includes('부재')) {
                        assemblyColIndex = i;
                        console.log(`부재명 열 발견: ${i+1}번째 열`);
                    } else if (cellStr.includes('생산') && cellStr.includes('잔량')) {
                        productionQuantityIndex = i;
                        console.log(`생산잔량 열 발견: ${i+1}번째 열`);
                    }
                }
            }
        }
    }
    
    console.log(`사용할 열 인덱스 - 부재명: ${assemblyColIndex+1}번째 열, 생산잔량: ${productionQuantityIndex+1}번째 열`);
    
    // 제외할 키워드 목록 (소계/합계만 제외)
    const excludeKeywords = ['소계', '합계', 'total', 'subtotal'];
    
    // 중복 체크를 위한 집합
    const processedItems = new Set();
    
    // 데이터 처리 (데이터 시작 행부터)
    for (let i = dataStartIndex; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || !Array.isArray(row)) {
            continue;
        }
        
        // 부재명과 수량 추출
        const assemblyStr = row[assemblyColIndex] ? String(row[assemblyColIndex]).trim() : '';
        let quantityStr = row[productionQuantityIndex] ? String(row[productionQuantityIndex]).trim() : '0';
        
        // 비어있는 행 건너뛰기
        if (!assemblyStr) {
            continue;
        }
        
        // 소계, 합계 등 제외 (대소문자 구분 없이)
        const assemblyLower = assemblyStr.toLowerCase();
        const shouldExclude = excludeKeywords.some(keyword => 
            assemblyLower.includes(keyword.toLowerCase())
        );
        
        if (shouldExclude) {
            console.log(`제외된 행 ${i+1} (제외 키워드):`, assemblyStr);
            continue;
        }
        
        // 수량 변환
        let quantity = 0;
        try {
            // 숫자만 추출 (쉼표 제거)
            quantityStr = quantityStr.replace(/,/g, '');
            quantity = parseInt(quantityStr, 10);
        } catch (e) {
            console.log(`행 ${i+1} 제외: 수량 변환 오류 - ${quantityStr}`);
            continue;
        }
        
        // 유효하지 않은 수량 건너뛰기
        if (isNaN(quantity) || quantity <= 0) {
            console.log(`행 ${i+1} 제외: 유효하지 않은 수량 - ${quantity}`);
            continue;
        }
        
        // 중복 체크 (같은 날짜+부재번호 조합 체크)
        const itemKey = `${finalDate}-${assemblyStr}`;
        if (processedItems.has(itemKey)) {
            console.log(`행 ${i+1} 제외: 중복된 데이터 - ${itemKey}`);
            continue;
        }
        
        // 중복 처리를 위해 키 추가
        processedItems.add(itemKey);
        
        // 유효한 데이터 추가
        const item = {
            '부재명': assemblyStr,
            '수량': quantity,
            '날짜': finalDate
        };
        
        results.push({
            date: finalDate,
            assemblyNumber: assemblyStr,
            quantity: quantity,
            company: 'jinsungpc'
        });
        
        console.log(`행 ${i+1} 데이터 추가:`, item);
    }
    
    console.log(`[${finalDate}] 처리된 데이터: ${results.length} 건`);
    if (results.length > 0) {
        console.log('첫 번째 데이터:', results[0]);
        console.log('마지막 데이터:', results[results.length - 1]);
    }
    
    return results;
}

// 문서 내에서 날짜 찾기 함수
function findDateInDocument(jsonData) {
    // 날짜 패턴: YYYY-MM-DD, YYYY.MM.DD, YYYY/MM/DD, YYYYMMDD 형식 등
    const datePatterns = [
        /\b(20\d{2})[-\.\/]?(0[1-9]|1[0-2])[-\.\/]?(0[1-9]|[12]\d|3[01])\b/g,  // YYYY-MM-DD, YYYY.MM.DD, YYYY/MM/DD
        /\b(0[1-9]|1[0-2])[-\.\/]?(0[1-9]|[12]\d|3[01])[-\.\/]?(20\d{2})\b/g,   // MM-DD-YYYY, MM.DD.YYYY, MM/DD/YYYY
        /\b(20\d{2}년)\s*([0-9]{1,2}월)\s*([0-9]{1,2}일)\b/g  // 한글 날짜 (YYYY년 MM월 DD일)
    ];
    
    // 첫 50행 내에서 날짜 찾기 (더 많은 행 검색)
    for (let i = 0; i < Math.min(50, jsonData.length); i++) {
        const row = jsonData[i];
        if (!row || !Array.isArray(row)) continue;
        
        // 행의 각 셀에서 날짜 찾기
        for (let j = 0; j < row.length; j++) {
            const cell = row[j];
            if (!cell) continue;
            
            const cellStr = String(cell);
            console.log(`행 ${i+1}, 열 ${j+1} 검사 중 - 셀 내용: "${cellStr}"`);
            
            // 날짜가 포함된 특수 제목 셀 검사
            if (cellStr.includes('생산일보') || cellStr.includes('생산 실적')) {
                console.log(`생산일보 제목 셀 발견: "${cellStr}"`);
                
                // 제목에서 날짜 찾기 (예: "2025-03-19 생산일보")
                const dateMatch = cellStr.match(/(20\d{2})[-\.\/\s]([01]?\d)[-\.\/\s]([0-3]?\d)/);
                if (dateMatch) {
                    const year = dateMatch[1];
                    const month = String(parseInt(dateMatch[2])).padStart(2, '0');
                    const day = String(parseInt(dateMatch[3])).padStart(2, '0');
                    console.log(`제목에서 날짜 발견: ${year}년 ${month}월 ${day}일`);
                    return `${year}${month}${day}`;
                }
            }
            
            // 한글 날짜 패턴 특별 처리 (예: 2023년 3월 5일)
            const koreanDateMatch = cellStr.match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/);
            if (koreanDateMatch) {
                const year = koreanDateMatch[1];
                const month = String(parseInt(koreanDateMatch[2])).padStart(2, '0');
                const day = String(parseInt(koreanDateMatch[3])).padStart(2, '0');
                console.log(`한글 날짜 발견: ${year}년 ${month}월 ${day}일`);
                return `${year}${month}${day}`;
            }
            
            // 각 패턴으로 검색
            for (let pattern of datePatterns) {
                const matches = cellStr.match(pattern);
                if (matches && matches.length > 0) {
                    let match = matches[0];
                    console.log(`날짜 패턴 발견: ${match}`);
                    
                    // 형식에 따라 YYYYMMDD로 변환
                    let year, month, day;
                    
                    if (match.length === 8 && /^\d{8}$/.test(match)) {
                        // 이미 YYYYMMDD 형식
                        console.log(`YYYYMMDD 형식 날짜 발견: ${match}`);
                        return match;
                    } else if (match.includes('-') || match.includes('.') || match.includes('/')) {
                        // YYYY-MM-DD 또는 MM-DD-YYYY 형식
                        const parts = match.split(/[-\.\/]/);
                        
                        if (parts.length === 3) {
                            if (parts[0].length === 4) {
                                // YYYY-MM-DD
                                year = parts[0];
                                month = parts[1].padStart(2, '0');
                                day = parts[2].padStart(2, '0');
                            } else {
                                // MM-DD-YYYY
                                year = parts[2];
                                month = parts[0].padStart(2, '0');
                                day = parts[1].padStart(2, '0');
                            }
                            
                            const formattedDate = `${year}${month}${day}`;
                            console.log(`변환된 날짜: ${formattedDate} (원본: ${match})`);
                            return formattedDate;
                        }
                    }
                }
            }
        }
    }
    
    // 날짜를 못 찾은 경우 null 반환
    console.log('문서에서 날짜를 찾을 수 없습니다.');
    return null;
}

module.exports = {
    parseJinsungPCData
}; 