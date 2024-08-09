const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const sheets = spreadsheet.getSheets();

function updateSheetWithWeightedScores() {
    sheets.forEach(sheet => {
        if (sheet.getName().includes("(설문)스프린트평가")) {
            const lastRow = sheet.getLastRow(); // 데이터가 있는 마지막 행 번호 가져오기

            for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
                let totalWeightedScore = calculateRowWeightedScore(sheet, rowIndex);
                sheet.getRange(rowIndex, 7).setValue(totalWeightedScore.toFixed(1));
            }
        }
    });
}

function calculateRowWeightedScore(sheet,rowIndex) {
    const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let itemCount = [0, 0, 0, 0, 0, 0]; // 카테고리별 문항수 초기화

    // 카테고리별 문항수 계산
    let categoryLine = ['[A.', '[B.', '[C.', '[D.', '[E.', '[F.'];
    firstRow.forEach(cellContent => {
        let cell = cellContent.toString(); // 셀 내용을 문자열로 변환
        for (let i = 0; i < 6; i++) {
            if (cell.startsWith(categoryLine[i])) {
                itemCount[i]++;
                break;
            }
        }
    });

    const weightedScore = [50, 20, 10, 10, 10]; // 가중치 점수
    let receivedScore = [0, 0, 0, 0, 0, 0]; // 카테고리별 받은 점수의 합계
    const itemSpecificScoring = [1, 6, 6, 6, 6, 1]; // 카테고리별 최대 척도
    let categoryMaxScore = []; // 카테고리별 최대 합계 점수

    for (let i = 0; i < 5; i++) {
        categoryMaxScore.push(itemCount[i] * itemSpecificScoring[i]); // 카테고리별 최대 점수 = 문항수 * 척도
    }

    // 카테고리별 받은 점수 계산
    const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    firstRow.forEach((cellContent, index) => {
        let cell = (cellContent || '').toString(); // null이나 undefined를 빈 문자열로 처리 후 문자열 변환
        for (let i = 0; i < 6; i++) {
            if (cell.startsWith(categoryLine[i])) {
                receivedScore[i] = (receivedScore[i] || 0) + (rowData[index] || 0);
            }
        }
    });

    let bonusPoint = 0; // 가중치 계산
    if (receivedScore[5] <= 5) {
        bonusPoint = receivedScore[5];
    } else if (receivedScore[5] > 5) {
        bonusPoint = 5;
    }

    // 가중치 적용된 카테고리별 점수 = (가중치 점수 * 받은 점수) / 카테고리별 만점 점수
    let weightedCategoryScore = [];
    for (let i = 0; i < 5; i++) {
        weightedCategoryScore.push((weightedScore[i] * receivedScore[i]) / categoryMaxScore[i]);
    }

    let totalWeightedScore = 0;
    for (let i = 0; i < 5; i++) {
        if (isNaN(weightedCategoryScore[i])) continue;
        totalWeightedScore += weightedCategoryScore[i];
    }

    totalWeightedScore += bonusPoint;

    // H열에 총점 저장 ('H열' = 8)
    return totalWeightedScore;
}