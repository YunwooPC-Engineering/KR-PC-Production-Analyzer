// 원자재 데이터 관리 클래스
class MaterialsManager {
    constructor() {
        this.materials = [];
        this.initializeEventListeners();
    }

    // 이벤트 리스너 초기화
    initializeEventListeners() {
        const form = document.getElementById('materialForm');
        const viewByMonthBtn = document.getElementById('viewByMonth');
        const exportExcelBtn = document.getElementById('exportExcel');

        if (form) {
            form.addEventListener('submit', (e) => {
                e.preventDefault();
                this.handleFileUpload();
            });
        }

        if (viewByMonthBtn) {
            viewByMonthBtn.addEventListener('click', () => this.showMonthlyView());
        }

        if (exportExcelBtn) {
            exportExcelBtn.addEventListener('click', () => this.exportToExcel());
        }

        // 페이지 로드 시 저장된 데이터 불러오기
        this.loadSavedData();
    }

    // 파일 업로드 처리
    async handleFileUpload() {
        const fileInput = document.getElementById('materialFile');
        const processDate = document.getElementById('processDate');
        
        if (!fileInput.files.length) {
            alert('파일을 선택해주세요.');
            return;
        }

        if (!processDate.value) {
            alert('처리 날짜를 선택해주세요.');
            return;
        }

        const file = fileInput.files[0];
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // 첫 번째 시트 데이터 파싱
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                
                this.processExcelData(jsonData, processDate.value);
            } catch (error) {
                console.error('파일 처리 중 오류 발생:', error);
                alert('파일 처리 중 오류가 발생했습니다.');
            }
        };

        reader.readAsArrayBuffer(file);
    }

    // 엑셀 데이터 처리
    processExcelData(data, processDate) {
        // 헤더 행 건너뛰기
        const rows = data.slice(1);
        
        const newMaterials = rows.map(row => {
            // 데이터가 비어있거나 유효하지 않은 경우 건너뛰기
            if (!row || row.length < 5) return null;

            return {
                date: processDate,
                code: row[0] || '',
                name: row[1] || '',
                specification: row[2] || '',
                unit: row[3] || '',
                quantity: parseFloat(row[4]) || 0,
                price: parseFloat(row[5]) || 0,
                amount: parseFloat(row[6]) || 0
            };
        }).filter(item => item !== null);

        // 기존 데이터와 병합
        this.materials = [...this.materials, ...newMaterials];
        
        // 데이터 저장 및 화면 갱신
        this.saveData();
        this.updateTable();
        
        // 폼 초기화
        document.getElementById('materialForm').reset();
    }

    // 데이터 저장
    saveData() {
        localStorage.setItem('materialsData', JSON.stringify(this.materials));
    }

    // 저장된 데이터 불러오기
    loadSavedData() {
        const savedData = localStorage.getItem('materialsData');
        if (savedData) {
            this.materials = JSON.parse(savedData);
            this.updateTable();
        }
    }

    // 테이블 업데이트
    updateTable(data = this.materials) {
        const tbody = document.querySelector('#materialsTable tbody');
        if (!tbody) return;

        tbody.innerHTML = '';
        
        data.forEach(item => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${this.formatDate(item.date)}</td>
                <td>${item.code}</td>
                <td>${item.name}</td>
                <td>${item.specification}</td>
                <td>${item.unit}</td>
                <td>${item.quantity.toLocaleString()}</td>
                <td>${item.price.toLocaleString()}</td>
                <td>${item.amount.toLocaleString()}</td>
            `;
            tbody.appendChild(row);
        });
    }

    // 월별 보기
    showMonthlyView() {
        const monthlyData = {};
        
        // 데이터를 월별로 그룹화
        this.materials.forEach(item => {
            const month = item.date.substring(0, 7); // YYYY-MM 형식
            if (!monthlyData[month]) {
                monthlyData[month] = [];
            }
            monthlyData[month].push(item);
        });

        // 월별 합계 계산
        const summaryData = Object.entries(monthlyData).map(([month, items]) => {
            const totalQuantity = items.reduce((sum, item) => sum + item.quantity, 0);
            const totalAmount = items.reduce((sum, item) => sum + item.amount, 0);

            return {
                date: month,
                code: '월별 합계',
                name: '',
                specification: '',
                unit: '',
                quantity: totalQuantity,
                price: 0,
                amount: totalAmount
            };
        });

        // 월별 정렬
        summaryData.sort((a, b) => b.date.localeCompare(a.date));
        
        // 테이블 업데이트
        this.updateTable(summaryData);
    }

    // Excel 내보내기
    exportToExcel() {
        // 데이터를 월별로 그룹화
        const monthlyData = {};
        this.materials.forEach(item => {
            const month = item.date.substring(0, 7);
            if (!monthlyData[month]) {
                monthlyData[month] = [];
            }
            monthlyData[month].push(item);
        });

        // 워크북 생성
        const wb = XLSX.utils.book_new();
        
        // 전체 데이터 시트
        const allDataWS = XLSX.utils.json_to_sheet(this.materials);
        XLSX.utils.book_append_sheet(wb, allDataWS, '전체 데이터');

        // 월별 시트 생성
        Object.entries(monthlyData).forEach(([month, items]) => {
            const ws = XLSX.utils.json_to_sheet(items);
            XLSX.utils.book_append_sheet(wb, ws, month);
        });

        // 파일 저장
        XLSX.writeFile(wb, '원자재입고현황.xlsx');
    }

    // 날짜 포맷팅
    formatDate(dateStr) {
        const date = new Date(dateStr);
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }
}

// 페이지 로드 시 MaterialsManager 인스턴스 생성
document.addEventListener('DOMContentLoaded', () => {
    window.materialsManager = new MaterialsManager();
}); 