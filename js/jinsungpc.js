// 진성피씨 업체 분석용 스크립트

// 진성피씨 엑셀 데이터 구조 분석 및 처리
function parseJinsungPCData(jsonData, date) {
    const processedData = [];
    
    console.log('진성피씨 데이터 파싱 시작:', date);
    console.log('데이터 행 수:', jsonData.length);
    
    try {
        // 데이터 행 위치 찾기
        let foundHeader = false;
        let assemblyColIndex = null;
        let quantityColIndex = null;
        
        // 먼저 헤더 행을 찾습니다
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length === 0) continue;
            
            // 각 셀을 검사하여 '부재명'을 찾습니다
            for (let j = 0; j < row.length; j++) {
                const cellValue = row[j] ? row[j].toString().trim() : '';
                if (cellValue === '부재명') {
                    foundHeader = true;
                    assemblyColIndex = j;
                    console.log(`헤더 행 발견: ${i+1}행, 부재명 열: ${j+1}열`);
                    
                    // 생산수량(EA) 열 찾기
                    for (let k = 0; k < row.length; k++) {
                        const quantityCellValue = row[k] ? row[k].toString().trim() : '';
                        if (quantityCellValue === '생산수량(EA)') {
                            quantityColIndex = k;
                            console.log(`생산수량(EA) 열 발견: ${k+1}열`);
                            break;
                        }
                    }
                    break;
                }
            }
            
            if (foundHeader) {
                break;
            }
        }
        
        if (!foundHeader || assemblyColIndex === null || quantityColIndex === null) {
            console.warn('헤더를 찾을 수 없거나 필요한 열을 찾지 못했습니다.');
            console.log('데이터 형식 확인 필요:', jsonData.slice(0, 10));
            return processedData;
        }
        
        // 각 행을 순회하면서 데이터 처리
        let dataStarted = false;
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length <= Math.max(assemblyColIndex, quantityColIndex)) {
                continue; // 열 수가 부족하면 건너뜁니다
            }
            
            // 헤더 행을 찾은 후의 데이터만 처리
            if (!dataStarted) {
                const cellValue = row[assemblyColIndex] ? row[assemblyColIndex].toString().trim() : '';
                if (cellValue === '부재명') {
                    dataStarted = true;
                    continue;
                }
            }
            
            if (!dataStarted) continue;
            
            // 부재명과 생산수량 확인
            const assemblyNumber = row[assemblyColIndex];
            const quantity = row[quantityColIndex];
            
            if (!assemblyNumber) {
                continue; // 부재명이 없으면 건너뜁니다
            }
            
            const assemblyStr = assemblyNumber.toString().trim();
            
            // 소계, 합계 등 제외
            if (assemblyStr.toLowerCase().includes('소계') || 
                assemblyStr.toLowerCase().includes('합계') || 
                assemblyStr.toLowerCase().includes('total') || 
                assemblyStr.toLowerCase().includes('subtotal') ||
                assemblyStr.toLowerCase().includes('페이지')) {
                continue;
            }
            
            // 데이터 행 처리 - 생산수량이 유효한 숫자이고 0이 아닌 경우만 추가
            if (quantity && !isNaN(Number(quantity)) && Number(quantity) > 0) {
                console.log(`행 ${i+1}: 데이터 추가 - 부재명: ${assemblyStr}, 수량: ${quantity}`);
                processedData.push({
                    date: date,
                    assemblyNumber: assemblyStr,
                    quantity: Number(quantity),
                    company: 'jinsungpc'
                });
            }
        }
        
        console.log(`[${date}] 처리된 데이터:`, processedData.length, '건');
        if (processedData.length > 0) {
            console.log('첫 번째 데이터:', processedData[0]);
            console.log('마지막 데이터:', processedData[processedData.length - 1]);
        } else {
            console.warn('처리된 데이터가 없습니다!');
        }
        
        return processedData;
    } catch (error) {
        console.error('진성피씨 데이터 파싱 중 오류:', error);
        throw error;
    }
} 