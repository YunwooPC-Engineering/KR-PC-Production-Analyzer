// 이수이앤씨 음성공장 생산일보 분석 (웹 버전)
class IsueDataParser {
    constructor() {
        this.currentYear = new Date().getFullYear();
        this.yearFromPath = null;
        // 년도 추출 패턴들
        this.yearPatterns = {
            // 파일 경로에서 추출하는 패턴
            path: [
                /(\d{2})년/, // "25년" -> 2025년
                /20(\d{2})년/, // "2025년" -> 2025년
                /[\/\\](\d{2})[\/\\]/, // 폴더구조에서 "/25/" -> 2025년
                /[\/\\]20(\d{2})[\/\\]/ // 폴더구조에서 "/2025/" -> 2025년
            ],
            // 파일명에서 추출하는 패턴
            filename: [
                /20(\d{2})[-_]?(\d{2})[-_]?(\d{2})/, // "20250321" 또는 "2025-03-21"
                /(\d{2})[-_]?(\d{2})[-_]?(\d{2})/, // "250321" 또는 "25-03-21"
                /(\d{2})(\d{2})(\d{2})/ // "250321"
            ],
            // 파일 내용에서 추출하는 패턴
            content: [
                /20(\d{2})년\s*(\d{1,2})월\s*(\d{1,2})일/, // "2025년 3월 21일"
                /(\d{2})년\s*(\d{1,2})월\s*(\d{1,2})일/, // "25년 3월 21일"
                /(\d{1,2})월\s*(\d{1,2})일\s*\(?(\d{4})\)?/, // "3월 21일 (2025)"
                /(\d{4})[-\.\/](\d{1,2})[-\.\/](\d{1,2})/, // "2025-03-21"
                /(\d{1,2})[-\.\/](\d{1,2})[-\.\/](\d{4})/, // "21-03-2025"
            ]
        };
    }

    // 날짜 형식 변환 (MMDD -> YYYYMMDD)
    formatDate(mmdd) {
        const year = this.yearFromPath || this.currentYear;
        return `${year}${mmdd}`;
    }

    // 부재번호 유효성 검사
    isValidAssemblyNumber(assemblyNumber) {
        if (!assemblyNumber) return false;
        
        // 소계, 합계 등의 키워드가 포함된 행은 제외
        if (this.isExcludedKeyword(assemblyNumber)) {
            return false;
        }

        // 부재명 형식 검사: 30-100-0100 또는 R41-120-0100 형식
        const pattern = /^[A-Za-z]?\d{2}-\d{3}-\d{4}$/;
        return pattern.test(String(assemblyNumber));
    }

    isExcludedKeyword(text) {
        if (!text || typeof text !== 'string') return true;
        const excludeKeywords = ['소계', '합계', 'total', 'subtotal'];
        const lowerText = text.toLowerCase();
        return excludeKeywords.some(keyword => lowerText.includes(keyword.toLowerCase()));
    }

    // 여러 소스에서 년도 정보 추출 시도
    extractYearFromMultipleSources(file, jsonData) {
        // 1. 파일 경로에서 추출 시도
        const pathYear = this.extractYearFromPath(file.name || file.path || '');
        if (pathYear) {
            console.log(`파일 경로에서 연도 추출: ${pathYear}`);
            return pathYear;
        }
        
        // 2. 파일명에서 추출 시도
        const filenameYear = this.extractYearFromFilename(file.name);
        if (filenameYear) {
            console.log(`파일명에서 연도 추출: ${filenameYear}`);
            return filenameYear;
        }
        
        // 3. 파일 내용에서 추출 시도
        if (jsonData && Array.isArray(jsonData)) {
            const contentYear = this.extractYearFromContent(jsonData);
            if (contentYear) {
                console.log(`파일 내용에서 연도 추출: ${contentYear}`);
                return contentYear;
            }
        }
        
        // 4. 파일 메타데이터에서 추출 시도 (파일 수정일/생성일)
        if (file.lastModified) {
            const lastModDate = new Date(file.lastModified);
            const metaYear = String(lastModDate.getFullYear());
            console.log(`파일 메타데이터에서 연도 추출: ${metaYear}`);
            return metaYear;
        }
        
        // 5. 현재 연도를 기본값으로 사용
        console.log(`연도 정보를 찾을 수 없어 현재 연도 사용: ${this.currentYear}`);
        return String(this.currentYear);
    }

    // 파일 경로에서 연도 정보 추출
    extractYearFromPath(filePath) {
        if (!filePath) return null;
        
        for (const pattern of this.yearPatterns.path) {
            const match = filePath.match(pattern);
            if (match && match[1]) {
                // 2자리 연도인 경우 앞에 20 붙이기
                const yearStr = match[1].length === 2 ? '20' + match[1] : match[1];
                // 유효한 연도 범위 확인 (2000-2099)
                const year = parseInt(yearStr, 10);
                if (year >= 2000 && year <= 2099) {
                    this.yearFromPath = yearStr;
                    return this.yearFromPath;
                }
            }
        }
        
        return null;
    }

    // 파일명에서 연도 추출
    extractYearFromFilename(filename) {
        if (!filename) return null;
        
        for (const pattern of this.yearPatterns.filename) {
            const match = filename.match(pattern);
            if (match) {
                // 첫 번째 그룹이 연도인 패턴
                const yearStr = match[1].length === 2 ? '20' + match[1] : match[1];
                const year = parseInt(yearStr, 10);
                if (year >= 2000 && year <= 2099) {
                    return yearStr;
                }
            }
        }
        
        return null;
    }
    
    // 파일 내용에서 날짜 추출
    extractYearFromContent(jsonData) {
        // 처음 10행 내에서 날짜 정보 검색
        const maxRows = Math.min(10, jsonData.length);
        
        for (let i = 0; i < maxRows; i++) {
            const row = jsonData[i];
            if (!row) continue;
            
            // 행의 각 셀 검사
            for (let j = 0; j < row.length; j++) {
                const cellValue = String(row[j] || '');
                
                for (const pattern of this.yearPatterns.content) {
                    const match = cellValue.match(pattern);
                    if (match) {
                        // 패턴에 따라 연도 그룹 위치가 다를 수 있음
                        let yearStr;
                        if (pattern.toString().includes('(\d{4})')) {
                            // 4자리 연도 패턴
                            for (let g = 1; g <= match.length; g++) {
                                if (match[g] && match[g].length === 4) {
                                    yearStr = match[g];
                                    break;
                                }
                            }
                        } else if (pattern.toString().includes('(\d{2})년')) {
                            // 2자리 연도 + '년' 패턴
                            yearStr = '20' + match[1];
                        } else {
                            // 기타 패턴은 첫 번째 그룹이 연도
                            yearStr = match[1].length === 2 ? '20' + match[1] : match[1];
                        }
                        
                        // 유효한 연도 범위 확인
                        if (yearStr) {
                            const year = parseInt(yearStr, 10);
                            if (year >= 2000 && year <= 2099) {
                                return yearStr;
                            }
                        }
                    }
                }
            }
        }
        
        return null;
    }

    // 파일명에서 월/일 추출
    extractDateFromFilename(filename) {
        if (!filename) return null;
        
        // 0321 형식 추출 (월/일만 있는 형식)
        const matchMMDD = filename.match(/(\d{2})(\d{2})/);
        if (matchMMDD) {
            const [_, month, day] = matchMMDD;
            // 유효한 월/일 범위 확인
            const monthNum = parseInt(month, 10);
            const dayNum = parseInt(day, 10);
            if (monthNum >= 1 && monthNum <= 12 && dayNum >= 1 && dayNum <= 31) {
                const year = this.yearFromPath || this.currentYear;
                return `${year}${month}${day}`;
            }
        }
        
        // 20250321 형식 추출 (연/월/일 모두 있는 형식)
        const matchYYYYMMDD = filename.match(/20(\d{2})(\d{2})(\d{2})/);
        if (matchYYYYMMDD) {
            const [_, yy, month, day] = matchYYYYMMDD;
            const monthNum = parseInt(month, 10);
            const dayNum = parseInt(day, 10);
            if (monthNum >= 1 && monthNum <= 12 && dayNum >= 1 && dayNum <= 31) {
                return `20${yy}${month}${day}`;
            }
        }
        
        // 다른 형식 추출 시도 (25-03-21 등)
        const matchYYMMDD = filename.match(/(\d{2})[-_.]?(\d{2})[-_.]?(\d{2})/);
        if (matchYYMMDD) {
            const [_, yy, month, day] = matchYYMMDD;
            const monthNum = parseInt(month, 10);
            const dayNum = parseInt(day, 10);
            if (monthNum >= 1 && monthNum <= 12 && dayNum >= 1 && dayNum <= 31) {
                return `20${yy}${month}${day}`;
            }
        }
        
        return null;
    }

    // 헤더 행 찾기 - 이수이앤씨 음성공장 형식에 맞게 개선
    findHeaderRows(jsonData) {
        // 헤더 범위를 찾기 (보통 여러 행에 걸쳐 병합된 헤더가 있음)
        // 최대 10행까지 탐색
        const maxSearchRows = Math.min(10, jsonData.length);
        
        // 생산, 수량(매), 금일 등의 키워드가 있는 행 찾기
        let productionRow = -1;      // '생산' 키워드가 있는 행
        let quantityRow = -1;        // '수량(매)' 키워드가 있는 행
        let todayRow = -1;           // '금일' 키워드가 있는 행
        
        // 부재번호 열 인덱스 (일반적으로 고정)
        let assemblyColIndex = -1;
        
        // 각 키워드의 열 위치
        let productionColIndex = -1;
        let quantityColIndex = -1;
        let todayColIndex = -1;
        
        // 헤더 행들 탐색
        for (let i = 0; i < maxSearchRows; i++) {
            const row = jsonData[i];
            if (!row) continue;
            
            for (let j = 0; j < row.length; j++) {
                if (!row[j]) continue;
                
                const cellValue = String(row[j]).toLowerCase();
                
                // 부재번호 열 찾기
                if (cellValue.includes('부재') && assemblyColIndex === -1) {
                    assemblyColIndex = j;
                }
                
                // '생산' 키워드 찾기
                if (cellValue === '생산' && productionRow === -1) {
                    productionRow = i;
                    productionColIndex = j;
                }
                
                // '수량(매)' 또는 '수량' 키워드 찾기
                if ((cellValue.includes('수량') && cellValue.includes('매') || 
                     cellValue === '수량') && quantityRow === -1) {
                    quantityRow = i;
                    quantityColIndex = j;
                }
                
                // '금일' 키워드 찾기
                if (cellValue === '금일' && todayRow === -1) {
                    todayRow = i;
                    todayColIndex = j;
                }
            }
        }
        
        console.log('헤더 구조 분석:');
        console.log('생산 행:', productionRow, '열:', productionColIndex);
        console.log('수량(매) 행:', quantityRow, '열:', quantityColIndex);
        console.log('금일 행:', todayRow, '열:', todayColIndex);
        console.log('부재번호 열:', assemblyColIndex);
        
        // 가장 마지막 헤더 행 결정 (금일, 수량, 생산 행 중 가장 큰 값)
        const lastHeaderRow = Math.max(productionRow, quantityRow, todayRow);
        
        // 금일 열 인덱스가 있으면 그 위치를 사용, 없으면 생산/수량 열의 위치 활용
        let targetQuantityColIndex = todayColIndex;
        if (targetQuantityColIndex === -1 && quantityColIndex !== -1) {
            targetQuantityColIndex = quantityColIndex;
        } else if (targetQuantityColIndex === -1 && productionColIndex !== -1) {
            targetQuantityColIndex = productionColIndex;
        }
        
        // 금일/생산/수량 열을 찾지 못한 경우 기본 처리
        if (targetQuantityColIndex === -1) {
            // 부재번호 열을 찾았다면 그 바로 옆 열을 수량으로 간주
            if (assemblyColIndex !== -1) {
                targetQuantityColIndex = assemblyColIndex + 1;
            } else {
                // 그마저도 없으면 3번째 열을 부재번호, 4번째 열을 수량으로 간주
                assemblyColIndex = 2;
                targetQuantityColIndex = 3;
            }
        }
        
        // 일반적으로 데이터는 마지막 헤더 행 다음 행부터 시작
        // 만약 헤더 구조를 전혀 찾지 못했다면 2번째 행을 헤더로 간주
        const dataStartRow = lastHeaderRow !== -1 ? lastHeaderRow + 1 : 2;
        
        return {
            headerEndRow: lastHeaderRow !== -1 ? lastHeaderRow : 1,
            dataStartRow: dataStartRow,
            assemblyColumnIndex: assemblyColIndex !== -1 ? assemblyColIndex : 2,
            quantityColumnIndex: targetQuantityColIndex !== -1 ? targetQuantityColIndex : 3
        };
    }

    async parseExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = async (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    // 결과를 저장할 배열
                    const results = [];
                    
                    // 다양한 소스에서 연도 정보 추출
                    const year = this.extractYearFromMultipleSources(file, jsonData);
                    this.yearFromPath = year; // 추출한 연도 저장
                    
                    // 파일명에서 날짜 추출 시도
                    let date = this.extractDateFromFilename(file.name);
                    
                    // 파일명에서 날짜를 추출하지 못한 경우, 파일 메타데이터 사용
                    if (!date && file.lastModified) {
                        const lastModDate = new Date(file.lastModified);
                        const year = lastModDate.getFullYear();
                        const month = String(lastModDate.getMonth() + 1).padStart(2, '0');
                        const day = String(lastModDate.getDate()).padStart(2, '0');
                        date = `${year}${month}${day}`;
                    }
                    
                    // 그래도 날짜가 없으면 현재 날짜 사용
                    if (!date) {
                        const now = new Date();
                        const year = now.getFullYear();
                        const month = String(now.getMonth() + 1).padStart(2, '0');
                        const day = String(now.getDate()).padStart(2, '0');
                        date = `${year}${month}${day}`;
                    }
                    
                    console.log(`최종 사용할 날짜: ${date} (YYYYMMDD 형식)`);

                    // 이수이앤씨 음성공장 양식에 맞게 헤더 구조 분석
                    const { 
                        headerEndRow,
                        dataStartRow, 
                        assemblyColumnIndex, 
                        quantityColumnIndex 
                    } = this.findHeaderRows(jsonData);
                    
                    console.log('헤더 분석 결과:');
                    console.log('헤더 마지막 행:', headerEndRow + 1); // 0-index를 1-index로 변환
                    console.log('데이터 시작 행:', dataStartRow + 1);
                    console.log('부재번호 열:', assemblyColumnIndex + 1);
                    console.log('생산수량(금일) 열:', quantityColumnIndex + 1);

                    // 중복 체크를 위한 집합
                    const processedItems = new Set();

                    // 데이터 행 처리 (헤더 다음 행부터)
                    for (let i = dataStartRow; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        if (!row || row.length <= Math.max(assemblyColumnIndex, quantityColumnIndex)) continue;

                        let assemblyNumber = row[assemblyColumnIndex];
                        if (assemblyNumber === undefined || assemblyNumber === null) continue;
                        
                        // 부재번호가 숫자인 경우 문자열로 변환
                        if (typeof assemblyNumber === 'number') {
                            assemblyNumber = String(assemblyNumber);
                        } else {
                            assemblyNumber = String(assemblyNumber).trim();
                        }
                        
                        // 비어있는 행 건너뛰기
                        if (!assemblyNumber) continue;
                            
                        let quantityValue = row[quantityColumnIndex];
                        let quantity = 0;
                        
                        if (typeof quantityValue === 'number') {
                            quantity = quantityValue;
                        } else if (quantityValue) {
                            let quantityStr = String(quantityValue).trim();
                            // 쉼표 제거하고 숫자 변환
                            quantityStr = quantityStr.replace(/,/g, '');
                            quantity = parseFloat(quantityStr);
                        }

                        // 유효하지 않은 부재번호 또는 수량이 0 이하인 경우 스킵
                        if (!this.isValidAssemblyNumber(assemblyNumber) || isNaN(quantity) || quantity <= 0) {
                            continue;
                        }

                        // 중복 체크
                        const itemKey = `${date}-${assemblyNumber}`;
                        if (processedItems.has(itemKey)) continue;

                        processedItems.add(itemKey);

                        results.push({
                            date: date,
                            assemblyNumber: assemblyNumber,
                            quantity: quantity,
                            company: 'isue'
                        });
                    }

                    console.log(`파일 ${file.name}에서 ${results.length}개의 데이터를 파싱했습니다.`);
                    resolve(results);
                } catch (error) {
                    console.error('파일 파싱 중 오류:', error);
                    reject(error);
                }
            };

            reader.onerror = () => reject(new Error('파일을 읽는 중 오류가 발생했습니다.'));
            reader.readAsArrayBuffer(file);
        });
    }
}

// 전역 객체로 내보내기
window.IsueDataParser = IsueDataParser; 