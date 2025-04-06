// 진성피씨 파싱 테스트 스크립트

// 테스트 파일 경로 설정
const TEST_FILES = [
    "C:\\Users\\agio9\\Downloads\\부산BGF물류\\부산BGF물류\\진성피씨\\진성 생산실적_20230319.xlsx",
    "C:\\Users\\agio9\\Downloads\\부산BGF물류\\부산BGF물류\\진성피씨\\진성 생산실적_20230318.xlsx",
    "C:\\Users\\agio9\\Downloads\\부산BGF물류\\부산BGF물류\\진성피씨\\진성 생산실적_20230317.xlsx"
];

// 파일명에서 날짜 추출 함수
function extractDateFromFilename(filename) {
    const match = filename.match(/(\d{8})/);
    if (!match) {
        throw new Error(`파일명에서 날짜를 찾을 수 없습니다: ${filename}`);
    }
    
    // 추출한 날짜의 유효성 검사
    const dateStr = match[1];
    const year = parseInt(dateStr.substring(0, 4));
    const month = parseInt(dateStr.substring(4, 6));
    const day = parseInt(dateStr.substring(6, 8));
    
    // 현재 날짜
    const now = new Date();
    const currentYear = now.getFullYear();
    
    // 미래 연도 조정
    if (year > currentYear) {
        console.warn(`미래 날짜 감지: ${dateStr}, 현재 연도로 조정합니다.`);
        const adjustedYear = currentYear;
        return `${adjustedYear}${dateStr.substring(4)}`;
    }
    
    return dateStr;
}

// 단일 파일 테스트 함수
async function testJinsungPCParser(filePath) {
    console.log(`\n=== 파일 테스트 시작: ${filePath} ===`);
    
    try {
        // 파일에서 날짜 추출
        const date = extractDateFromFilename(filePath);
        console.log('추출된 날짜:', date);
        
        // 파일 읽기
        const response = await fetch(filePath);
        const arrayBuffer = await response.arrayBuffer();
        
        // 엑셀 데이터 파싱
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        console.log('엑셀 시트 이름:', workbook.SheetNames[0]);
        console.log('전체 행 수:', jsonData.length);
        
        // 데이터 파싱 테스트
        const results = parseJinsungPCData(jsonData, date);
        
        console.log(`\n=== 테스트 결과 ===`);
        console.log('처리된 데이터 수:', results.length);
        if (results.length > 0) {
            console.log('첫 번째 항목:', results[0]);
            console.log('마지막 항목:', results[results.length - 1]);
        }
        
        return results;
    } catch (error) {
        console.error('테스트 실패:', error);
        throw error;
    }
}

// 전체 테스트 실행 함수
async function runTests() {
    console.log('진성피씨 파싱 테스트 시작');
    const allResults = [];
    
    for (const filePath of TEST_FILES) {
        try {
            const results = await testJinsungPCParser(filePath);
            allResults.push(...results);
        } catch (error) {
            console.error(`파일 처리 실패: ${filePath}`, error);
        }
    }
    
    console.log('\n=== 전체 테스트 결과 ===');
    console.log('총 처리된 데이터 수:', allResults.length);
    console.log('날짜별 데이터 수:');
    const dateGroups = {};
    allResults.forEach(item => {
        dateGroups[item.date] = (dateGroups[item.date] || 0) + 1;
    });
    console.table(dateGroups);
}

// 개발 환경에서만 테스트 버튼 표시
if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') {
    document.addEventListener('DOMContentLoaded', () => {
        const button = document.createElement('button');
        button.textContent = '진성피씨 파싱 테스트 실행';
        button.className = 'btn btn-warning mt-2';
        button.style.position = 'fixed';
        button.style.bottom = '10px';
        button.style.right = '10px';
        button.style.zIndex = '1000';
        button.onclick = runTests;
        document.body.appendChild(button);
    });
} 