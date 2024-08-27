// ppt 말고 이미지 형태임. 그런데 파싱이 잘 안됨.

$('#generate-image').click(function (e) {
    e.preventDefault();
    let newFile = $('#newFile')[0].files[0];
    let messageContainer = $('#message');
    let downloadSection = $('#download-section');
    let canvas = $('#reportCanvas')[0];
    let ctx = canvas.getContext('2d');
    let downloadLink = $('#download-link');
   
    if (!newFile) {
        messageContainer.text('엑셀 파일을 선택해 주세요.').addClass('error-message');
        return;
    }

    let reader = new FileReader();
    reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: 'array' });
        let sheetName = workbook.SheetNames[0];
        let sheet = workbook.Sheets[sheetName];
        let csvText = XLSX.utils.sheet_to_csv(sheet, { FS: "\t" });

        generateImage(csvText, ctx, function() {
            canvas.style.display = 'block';

            // 이미지 다운로드 링크 업데이트
            canvas.toBlob(function(blob) {
                let url = URL.createObjectURL(blob);
                downloadLink.attr('href', url);
                downloadSection.show();
            });
        });
    };
    reader.readAsArrayBuffer(newFile);
});

function generateImage(csvText, ctx, callback) {
    let rows = csvText.split('\n').map(row => row.split('\t'));
    let headers = rows[0];
    let data = rows.slice(1);

    // 기본 설정
    let canvasWidth = 1200;
    let rowHeight = 50;
    let canvasHeight = rowHeight * (data.length + 1) + 50;
    let colWidths = [150, 300, 150, 400, 200]; // 열 너비를 적절히 조정

    ctx.canvas.width = canvasWidth;
    ctx.canvas.height = canvasHeight;

    // 이미지 배경 설정
    ctx.fillStyle = "white";
    ctx.fillRect(0, 0, canvasWidth, canvasHeight);

    // 텍스트 스타일 설정
    ctx.fillStyle = "black";
    ctx.font = "16px Arial";
   
    // 테이블 그리기
    drawTable(ctx, headers, data, colWidths, rowHeight);

    callback();
}

function drawTable(ctx, headers, data, colWidths, rowHeight) {
    let x = 10;
    let y = 30;

    // 헤더 그리기
    headers.forEach((header, i) => {
        ctx.fillText(header, x, y);
        ctx.strokeRect(x, y - rowHeight, colWidths[i], rowHeight);
        x += colWidths[i];
    });

    y += rowHeight;

    // 데이터 그리기
    data.forEach(row => {
        x = 10;
        row.forEach((cell, i) => {
            drawTextInCell(ctx, cell, x, y, colWidths[i], rowHeight);
            ctx.strokeRect(x, y - rowHeight, colWidths[i], rowHeight);
            x += colWidths[i];
        });
        y += rowHeight;
    });
}

function drawTextInCell(ctx, text, x, y, maxWidth, rowHeight) {
    let words = text.split(' ');
    let line = '';
    let lineHeight = 20;
    let currentY = y - rowHeight + lineHeight;

    words.forEach(word => {
        let testLine = line + word + ' ';
        let metrics = ctx.measureText(testLine);
        let testWidth = metrics.width;

        if (testWidth > maxWidth && line !== '') {
            ctx.fillText(line, x, currentY);
            line = word + ' ';
            currentY += lineHeight;
        } else {
            line = testLine;
        }
    });

    ctx.fillText(line, x, currentY);
}

