// 이수이앤씨 음성공장 생산일보 분석
class IsueAnalyzer extends FactoryAnalyzer {
    constructor() {
        super('이수이앤씨 음성공장');
    }

    // 날짜 형식 변환 (MMDD -> YYYY-MM-DD)
    formatDate(dateString) {
        // 현재 연도 가져오기
        const currentYear = new Date().getFullYear();
        const year = currentYear;  // 기본값으로 현재 연도 사용
        
        if (dateString && dateString.length === 4) {
            const month = dateString.substring(0, 2);
            const day = dateString.substring(2);
            return `${year}-${month}-${day}`;
        }
        return dateString;
    }

    // 부재번호 유효성 검사
    isValidAssemblyNumber(assemblyNumber) {
        if (!assemblyNumber) return false;
        
        const str = assemblyNumber.toString().trim();
        
        // 소계, 합계 등 제외
        if (str.toLowerCase().includes('소계') || 
            str.toLowerCase().includes('합계') || 
            str.toLowerCase().includes('total') || 
            str.toLowerCase().includes('subtotal')) {
            return false;
        }
        
        // 숫자와 하이픈으로만 구성된 부재번호만 허용
        // 예: 123-456-7890 형식
        const pattern = /^\d{3}-\d{3}-\d{4}$/;
        return pattern.test(str);
    }

    parseFactoryData(jsonData, workbook) {
        const parsedData = [];
        let foundHeader = false;
        let quantityColIndex = null;
        
        // 파일명에서 날짜 추출 (MMDD 형식)
        const fileName = workbook.Props && workbook.Props.Title ? workbook.Props.Title : '';
        const dateMatch = fileName.match(/(\d{4})/);
        const date = dateMatch ? this.formatDate(dateMatch[1]) : '';
        
        // 데이터 행 처리
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            
            // 부재번호 헤더 찾기
            if (row['B'] === '부재번호') {
                foundHeader = true;
                continue;
            }
            
            // '생산' > '수량(매)' > '금일' 열 찾기
            if (row['B'] === '생산' && !quantityColIndex) {
                // 해당 행에서 '수량(매)' 찾기
                for (let col in row) {
                    if (row[col] === '수량(매)') {
                        // 다음 행에서 '금일' 열의 인덱스 찾기
                        const nextRow = jsonData[i + 1];
                        for (let nextCol in nextRow) {
                            if (nextRow[nextCol] === '금일') {
                                quantityColIndex = nextCol;
                                break;
                            }
                        }
                        break;
                    }
                }
                continue;
            }
            
            // 헤더를 찾은 후의 데이터만 처리
            if (foundHeader && quantityColIndex) {
                const assemblyNumber = row['B'];
                const quantity = row[quantityColIndex];
                
                // 부재번호 유효성 검사 및 수량이 있는 행만 처리
                if (this.isValidAssemblyNumber(assemblyNumber) && quantity) {
                    parsedData.push({
                        date: date,
                        assemblyNumber: assemblyNumber.toString().trim(),
                        quantity: parseInt(quantity) || 0
                    });
                }
            }
        }
        
        return parsedData;
    }

    async processFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', bookProps: true });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 'A' });

                    // 데이터 파싱 (workbook 전체를 전달하여 날짜 정보도 접근 가능하게 함)
                    const parsedData = this.parseFactoryData(jsonData, workbook);
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
}

// 이수이앤씨 음성공장 분석기 인스턴스 생성
const isueAnalyzer = new IsueAnalyzer();

// 이수이앤씨 음성공장 분석 함수
function analyzeIsueFiles() {
    return isueAnalyzer.analyzeFiles();
} 