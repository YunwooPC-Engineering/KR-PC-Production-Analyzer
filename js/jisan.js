// 지산개발 생산일보 분석
class JisanAnalyzer extends FactoryAnalyzer {
    constructor() {
        super('지산개발');
    }

    // 날짜 형식 변환 (YYYYMMDD -> YYYY-MM-DD)
    formatDate(dateString) {
        if (dateString.length === 8) {
            return `${dateString.substring(0, 4)}-${dateString.substring(4, 6)}-${dateString.substring(6)}`;
        }
        return dateString;
    }

    parseFactoryData(jsonData, date) {
        const parsedData = [];
        let foundProductNumberHeader = false;
        
        // 지산개발 생산일보는 부재번호가 B열, 수량이 D열에 있음
        const assemblyCol = 'B';  // 부재번호 열
        const quantityCol = 'D';  // 수량 열
        
        // 날짜 형식 변환
        const formattedDate = this.formatDate(date);
        
        // 데이터 행 처리
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            
            // '제품번호' 헤더를 찾음
            if (row[assemblyCol] === '제품번호') {
                foundProductNumberHeader = true;
                continue;
            }
            
            // '제품번호' 헤더를 찾은 후의 데이터만 처리
            if (foundProductNumberHeader) {
                const assemblyNumber = row[assemblyCol];
                const quantity = row[quantityCol];
                
                // 부재번호와 수량이 있는 행만 처리
                if (assemblyNumber && quantity && 
                    !assemblyNumber.toString().toLowerCase().includes('소계') &&
                    !assemblyNumber.toString().toLowerCase().includes('합계') &&
                    !assemblyNumber.toString().toLowerCase().includes('total') &&
                    !assemblyNumber.toString().toLowerCase().includes('subtotal')) {
                    
                    parsedData.push({
                        date: formattedDate,
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
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 'A' });

                    // 파일명에서 날짜 추출 (YYYYMMDD 형식)
                    const dateMatch = file.name.match(/(\d{8})/);
                    const date = dateMatch ? dateMatch[1] : '';

                    // 데이터 파싱
                    const parsedData = this.parseFactoryData(jsonData, date);
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

// 지산개발 분석기 인스턴스 생성
const jisanAnalyzer = new JisanAnalyzer();

// 지산개발 분석 함수
function analyzeJisanFiles() {
    return jisanAnalyzer.analyzeFiles();
} 