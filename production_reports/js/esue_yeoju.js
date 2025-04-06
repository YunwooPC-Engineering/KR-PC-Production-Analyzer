// 이수이앤씨 - 여주공장 생산일보 분석
class EsueYeojuDataParser {
    constructor() {
        this.factoryName = '이수이앤씨 - 여주공장';
        this.excludedKeywords = ['소계', '합계', 'total', 'subtotal'];
        this.validYearRange = {
            start: 2024,
            end: 2025
        };
    }

    // 엑셀 파일 분석
    async parseExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = event => {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // 파일 수정 날짜 추출
                    const fileModifiedDate = new Date(file.lastModified);
                    const modifiedDateStr = this.formatDateToYYYYMMDD(fileModifiedDate);
                    
                    // 파일명에서 날짜 추출
                    const filenameDate = this.extractDateFromFilename(file.name);
                    
                    console.log(`파일 ${file.name}에서 추출한 날짜: ${filenameDate}`);
                    console.log(`파일 수정 날짜: ${modifiedDateStr}`);
                    
                    // 시트명에서 날짜 추출
                    let sheetNameDate = null;
                    for (const sheetName of workbook.SheetNames) {
                        const extractedDate = this.extractDateFromText(sheetName);
                        if (extractedDate) {
                            sheetNameDate = extractedDate;
                            console.log(`시트명 ${sheetName}에서 추출한 날짜: ${sheetNameDate}`);
                            break;
                        }
                    }
                    
                    // 파일 데이터 파싱 (셀 내부 날짜를 찾아냄)
                    const result = this.parseFactoryData(workbook, filenameDate, modifiedDateStr, sheetNameDate);
                    resolve(result);
                } catch (error) {
                    console.error(`엑셀 파일 처리 중 오류 발생:`, error);
                    reject(error);
                }
            };
            reader.onerror = () => reject(new Error('파일 읽기 오류'));
            reader.readAsArrayBuffer(file);
        });
    }

    // Date 객체를 YYYYMMDD 문자열로 변환
    formatDateToYYYYMMDD(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}${month}${day}`;
    }

    // 파일 이름에서 날짜 추출
    extractDateFromFilename(filename) {
        let possibleDates = [];
        
        // 1. 경로에서 연도와 월 패턴 확인 ("25년 3월", "2025년 3월" 같은 패턴)
        const yearMonthPattern = /[\/\\](?:20)?(\d{2})[년]\s*(\d{1,2})[월][\/\\]/;
        const yearMonthMatch = filename.match(yearMonthPattern);
        
        if (yearMonthMatch) {
            const year = `20${yearMonthMatch[1]}`;
            const month = String(parseInt(yearMonthMatch[2])).padStart(2, '0');
            console.log(`파일 경로에서 연도/월 감지: ${year}년 ${month}월`);
            
            // 2. 파일명에서 일자 추출 (MMDD 또는 DD 패턴)
            const dayPattern = /[\/\\](\d{2})[\s_]/;
            const dayMatch = filename.match(dayPattern);
            
            if (dayMatch) {
                const day = dayMatch[1];
                const dateStr = `${year}${month}${day}`;
                console.log(`파일명에서 날짜 조합: ${dateStr}`);
                possibleDates.push({
                    date: dateStr,
                    confidence: 0.9, // 경로의 연도/월 + 파일명의 일자는 높은 신뢰도
                    source: '파일경로+파일명'
                });
            }
        }
        
        // 3. YYYYMMDD 패턴 확인
        const fullDatePattern = /(\d{4})[-_]?(\d{2})[-_]?(\d{2})/;
        const fullMatch = filename.match(fullDatePattern);
        if (fullMatch) {
            const dateStr = `${fullMatch[1]}${fullMatch[2]}${fullMatch[3]}`;
            console.log(`파일명에서 전체 날짜 발견: ${dateStr}`);
            possibleDates.push({
                date: dateStr,
                confidence: 0.8, // 전체 날짜 패턴은 비교적 높은 신뢰도
                source: '파일명'
            });
        }
        
        // 4. MMDD 패턴 확인
        const shortDatePattern = /(\d{2})(\d{2})/;
        const shortMatch = filename.match(shortDatePattern);
        if (shortMatch) {
            const month = parseInt(shortMatch[1]);
            const day = parseInt(shortMatch[2]);
            
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                // 현재 연도 사용
                const now = new Date();
                const dateStr = `${now.getFullYear()}${String(month).padStart(2, '0')}${String(day).padStart(2, '0')}`;
                console.log(`파일명에서 월/일 발견: ${dateStr}`);
                possibleDates.push({
                    date: dateStr,
                    confidence: 0.6, // MMDD 패턴은 낮은 신뢰도
                    source: '파일명(월일)'
                });
            }
        }
        
        // 가장 신뢰도 높은 날짜 반환
        if (possibleDates.length > 0) {
            possibleDates.sort((a, b) => b.confidence - a.confidence);
            console.log(`선택된 날짜: ${possibleDates[0].date} (신뢰도: ${possibleDates[0].confidence}, 출처: ${possibleDates[0].source})`);
            return possibleDates[0].date;
        }
        
        return null;
    }
    
    // 텍스트에서 날짜 추출 (문서명, 시트명 등에 사용)
    extractDateFromText(text) {
        if (!text) return null;
        
        let possibleDates = [];
        
        // 1. 년월일 완전한 패턴
        const fullPatterns = [
            /(\d{4})[-년\.\-]?\s*(\d{1,2})[-월\.\-]?\s*(\d{1,2})[-일]?/,
            /(\d{2})[-년\.\-]?\s*(\d{1,2})[-월\.\-]?\s*(\d{1,2})[-일]?/
        ];
        
        for (const pattern of fullPatterns) {
            const match = text.match(pattern);
            if (match) {
                let year = match[1];
                if (year.length === 2) year = `20${year}`;
                
                const month = String(parseInt(match[2])).padStart(2, '0');
                const day = String(parseInt(match[3])).padStart(2, '0');
                
                if (parseInt(month) >= 1 && parseInt(month) <= 12 && 
                    parseInt(day) >= 1 && parseInt(day) <= 31) {
                    const dateStr = `${year}${month}${day}`;
                    console.log(`텍스트에서 완전한 날짜 발견: ${dateStr}`);
                    possibleDates.push({
                        date: dateStr,
                        confidence: 0.85,
                        source: '텍스트(완전)'
                    });
                }
            }
        }
        
        // 2. 월일 패턴
        const monthDayPattern = /(\d{1,2})[-월\.\-]?\s*(\d{1,2})[-일]?/;
        const mdMatch = text.match(monthDayPattern);
        if (mdMatch) {
            const month = String(parseInt(mdMatch[1])).padStart(2, '0');
            const day = String(parseInt(mdMatch[2])).padStart(2, '0');
            
            if (parseInt(month) >= 1 && parseInt(month) <= 12 && 
                parseInt(day) >= 1 && parseInt(day) <= 31) {
                const now = new Date();
                const dateStr = `${now.getFullYear()}${month}${day}`;
                console.log(`텍스트에서 월/일 발견: ${dateStr}`);
                possibleDates.push({
                    date: dateStr,
                    confidence: 0.7,
                    source: '텍스트(월일)'
                });
            }
        }
        
        // 가장 신뢰도 높은 날짜 반환
        if (possibleDates.length > 0) {
            possibleDates.sort((a, b) => b.confidence - a.confidence);
            console.log(`선택된 날짜: ${possibleDates[0].date} (신뢰도: ${possibleDates[0].confidence}, 출처: ${possibleDates[0].source})`);
            return possibleDates[0].date;
        }
        
        return null;
    }
    
    // 문서 내용에서 날짜 추출
    extractDateFromDocument(jsonData) {
        if (!jsonData || jsonData.length === 0) return null;
        
        // 처음 10행 검색
        for (let i = 0; i < Math.min(10, jsonData.length); i++) {
            const row = jsonData[i];
            if (!row) continue;
            
            // 각 셀의 값을 문자열로 변환하여 검사
            for (const key in row) {
                const cellValue = String(row[key] || '');
                const extractedDate = this.extractDateFromText(cellValue);
                
                if (extractedDate) {
                    console.log(`문서에서 날짜 발견: ${extractedDate}, 셀: ${cellValue}`);
                    return extractedDate;
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

    // 여러 날짜 중 가장 적절한 날짜 선택
    selectMostRecentDate(dates) {
        // null, undefined 제거
        const validDates = dates.filter(date => date);
        
        if (validDates.length === 0) {
            // 날짜를 하나도 찾지 못한 경우 파일 메타데이터의 날짜 사용
            const now = new Date();
            const dateStr = this.formatDateToYYYYMMDD(now);
            console.log(`날짜를 찾을 수 없어 현재 날짜 사용: ${dateStr}`);
            return dateStr;
        }
        
        // 날짜 유효성 검증 및 정렬
        const validatedDates = validDates.map(date => {
            const year = parseInt(date.substring(0, 4));
            const month = parseInt(date.substring(4, 6));
            const day = parseInt(date.substring(6, 8));
            
            return {
                date,
                isValid: !isNaN(year) && !isNaN(month) && !isNaN(day) &&
                        month >= 1 && month <= 12 && day >= 1 && day <= 31,
                dateObj: new Date(year, month - 1, day)
            };
        }).filter(d => d.isValid);
        
        if (validatedDates.length === 0) {
            const now = new Date();
            const dateStr = this.formatDateToYYYYMMDD(now);
            console.log(`유효한 날짜가 없어 현재 날짜 사용: ${dateStr}`);
            return dateStr;
        }
        
        // 가장 최근 날짜 선택
        validatedDates.sort((a, b) => b.dateObj - a.dateObj);
        console.log(`선택된 최종 날짜: ${validatedDates[0].date}`);
        return validatedDates[0].date;
    }

    // 데이터 파싱
    parseFactoryData(workbook, filenameDate, modifiedDate, sheetNameDate) {
        const allParsedData = [];
        
        // 모든 시트를 처리
        for (const sheetName of workbook.SheetNames) {
            // 시트 이름 확인
            console.log(`시트 처리 중: ${sheetName}`);
            
            // 제외할 시트 이름
            if (sheetName.includes('사용금지') || 
                sheetName.includes('미사용') || 
                sheetName.includes('양식')) {
                console.log(`건너뛴 시트: ${sheetName} (제외 시트)`);
                continue;
            }
            
            const worksheet = workbook.Sheets[sheetName];
            
            try {
                // 엑셀 데이터를 JSON으로 변환
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                    header: 'A',
                    defval: '',
                    raw: false
                });
                
                console.log(`시트 ${sheetName}의 행 수: ${jsonData.length}`);
                
                // 빈 시트 건너뛰기
                if (!jsonData || jsonData.length < 3) {
                    console.log(`건너뛴 시트: ${sheetName} (데이터 부족)`);
                    continue;
                }
                
                // 문서 내에서 날짜 추출
                const documentDate = this.extractDateFromDocument(jsonData);
                
                // 가장 최신 날짜 선택
                const availableDates = [filenameDate, modifiedDate, sheetNameDate, documentDate];
                const finalDate = this.selectMostRecentDate(availableDates);
                
                console.log(`시트 ${sheetName}에 대한 날짜 후보:`, availableDates);
                console.log(`시트 ${sheetName}의 최종 선택 날짜: ${finalDate}`);
                
                // 이 시트의 데이터 파싱
                const sheetData = this.parseSheetData(jsonData, finalDate);
                
                // 유효한 데이터가 있는 경우만 추가
                if (sheetData && sheetData.length > 0) {
                    allParsedData.push(...sheetData);
                    console.log(`시트 ${sheetName} 처리 완료: ${sheetData.length}개 항목 추가됨`);
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
        
        // 처리된 데이터 샘플 로그
        if (finalData.length > 0) {
            console.log('처리된 데이터 샘플:');
            for (let i = 0; i < Math.min(5, finalData.length); i++) {
                console.log(`[${i+1}] ${finalData[i].date}, ${finalData[i].assemblyNumber}, ${finalData[i].quantity}`);
            }
        }
        
        return finalData;
    }
    
    // 시트 데이터 파싱
    parseSheetData(jsonData, sheetDate) {
        const parsedData = [];
        
        // 열 헤더와 데이터 시작행 찾기
        let assemblyColIndex = null;  // 부재번호 열
        let productionColIndex = null; // 생산량 열
        let headerRowIndex = -1;       // 헤더 행
        
        // 헤더 검색 (처음 20행 검색)
        for (let i = 0; i < Math.min(20, jsonData.length); i++) {
            const row = jsonData[i];
            if (!row) continue;
            
            let foundHeaders = false;
            
            for (const key in row) {
                const value = String(row[key] || '').trim().toLowerCase();
                
                // 부재번호 열 찾기
                if (value.includes('부재번호') || value.includes('품번') || 
                    value.includes('assy') || value.includes('자재') || 
                    value.includes('파트넘버') || value.includes('부품번호')) {
                    assemblyColIndex = key;
                    headerRowIndex = i;
                    foundHeaders = true;
                    console.log(`부재번호 열 발견: ${key}, 값: "${row[key]}"`);
                }
                
                // 생산량 열 찾기
                if (value.includes('생산') || value.includes('수량') || 
                    value.includes('생산잔량') || value.includes('생산량')) {
                    productionColIndex = key;
                    console.log(`생산량 열 발견: ${key}, 값: "${row[key]}"`);
                }
            }
            
            // 헤더를 찾았으면 루프 종료
            if (foundHeaders && assemblyColIndex && productionColIndex) {
                break;
            }
        }
        
        // 헤더를 찾지 못한 경우
        if (!assemblyColIndex || !productionColIndex) {
            console.log('자동 헤더 검색 실패, 데이터 구조 분석 시도...');
            
            // 모든 셀 데이터 확인 (처음 10행만)
            for (let i = 0; i < Math.min(10, jsonData.length); i++) {
                const row = jsonData[i];
                if (!row) continue;
                
                console.log(`행 ${i} 데이터:`, JSON.stringify(row));
            }
            
            // 부재번호 패턴으로 열 찾기
            for (let i = 0; i < Math.min(20, jsonData.length); i++) {
                const row = jsonData[i];
                if (!row) continue;
                
                for (const key in row) {
                    const value = String(row[key] || '');
                    
                    // 부재번호 패턴 검사 (XX-XXX-XXXX)
                    if (/^\d{2}-\d{3}-\d{4}$/.test(value)) {
                        assemblyColIndex = key;
                        headerRowIndex = i - 1; // 헤더는 이 행 바로 위
                        console.log(`부재번호 패턴 발견 (${i}행, ${key}열): ${value}`);
                        break;
                    }
                }
                
                if (assemblyColIndex) break;
            }
            
            // 부재번호 열을 찾았으면 생산량 열 추정
            if (assemblyColIndex) {
                const colCode = assemblyColIndex.charCodeAt(0);
                // 부재번호 열로부터 오른쪽으로 2~5칸 사이에 수량 열
                for (let i = 2; i <= 5; i++) {
                    const possibleQuantityCol = String.fromCharCode(colCode + i);
                    if (jsonData[headerRowIndex] && jsonData[headerRowIndex][possibleQuantityCol]) {
                        productionColIndex = possibleQuantityCol;
                        console.log(`추정된 생산량 열: ${productionColIndex}`);
                        break;
                    }
                }
            }
        }
        
        // 여전히 필요한 열을 찾지 못한 경우 기본값 사용
        if (!assemblyColIndex) {
            assemblyColIndex = 'B';  // 두 번째 열
            headerRowIndex = 1;      // 두 번째 행을 헤더로 가정
            console.log('부재번호 열을 찾을 수 없어 기본값 사용: B열');
        }
        
        if (!productionColIndex) {
            productionColIndex = 'E';  // 다섯 번째 열
            console.log('생산량 열을 찾을 수 없어 기본값 사용: E열');
        }
        
        if (headerRowIndex === -1) {
            headerRowIndex = 1; // 두 번째 행을 헤더로 가정
            console.log('헤더 행을 찾을 수 없어 기본값 사용: 1행');
        }
        
        console.log(`최종 설정 - 헤더 행: ${headerRowIndex}, 부재번호 열: ${assemblyColIndex}, 생산량 열: ${productionColIndex}`);
        
        // 데이터 행 처리 (헤더 다음 행부터)
        let processedCount = 0;
        let invalidCount = 0;
        let excludedCount = 0;
        
        for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;
            
            const assemblyNumber = String(row[assemblyColIndex] || '').trim();
            let productionQuantity = row[productionColIndex];
            
            // 디버깅용 로그
            if (i < headerRowIndex + 5) {
                console.log(`행 ${i}, 부재번호: "${assemblyNumber}", 생산량: "${productionQuantity}"`);
            }
            
            // 부재번호가 있고 제외 키워드가 아닌 경우만 처리
            if (assemblyNumber && !this.isExcludedRow(assemblyNumber)) {
                // 생산량 처리
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
                                const value = String(row[key] || '').replace(/,/g, '').trim();
                                if (value && !isNaN(parseFloat(value))) {
                                    const possibleQuantity = parseFloat(value);
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
                
                // 최종 데이터 추가 (수량이 있는 경우만)
                if (quantity > 0) {
                    parsedData.push({
                        // 기존 필드명 유지 (호환성)
                        date: sheetDate,
                        assemblyNumber: assemblyNumber,
                        quantity: quantity,
                        company: 'esue_yeoju',
                        
                        // 표준화된 필드명 추가
                        CompletedDate: sheetDate,
                        AssemblyNumber: assemblyNumber,
                        Quantity: quantity,
                        Company: 'esue_yeoju'
                    });
                    processedCount++;
                    
                    // 처리된 첫 5개 항목 출력 (디버깅용)
                    if (processedCount <= 5) {
                        console.log(`처리된 데이터 ${processedCount}: ${assemblyNumber}, 수량: ${quantity}`);
                    }
                } else {
                    invalidCount++;
                }
            } else if (assemblyNumber && this.isExcludedRow(assemblyNumber)) {
                excludedCount++;
            }
        }
        
        console.log(`시트 처리 완료 - 처리된 레코드: ${processedCount}, 무효한 수량: ${invalidCount}, 제외된 레코드: ${excludedCount}`);
        
        return parsedData;
    }
}

// 파서 인스턴스 생성을 위한 전역 접근점
window.EsueYeojuDataParser = EsueYeojuDataParser;