document.getElementById('generate-ppt').addEventListener('click', function() {
    let fileInput = document.getElementById('file-input');
    let messageContainer = document.getElementById('message');
   
    if (fileInput.files.length === 0) {
        messageContainer.textContent = 'Please select a CSV file to upload.';
        messageContainer.className = 'error-message';
        return;
    }

    let file = fileInput.files[0];
    let reader = new FileReader();
    reader.onload = function(e) {
        let text = e.target.result;
        generatePPT(text);
    };
    reader.readAsText(file);
});

function generatePPT(csvText) {
    let messageContainer = document.getElementById('message');

    let rows = csvText.split('\n').map(row => row.split('\t'));
    let headers = rows[0];
    let data = rows.slice(1);

    // Filter columns
    let columnsToKeep = ['고객불만번호', '제목', '불만발생일', '조사담당자의견', '고객요구사항'];
    let filteredHeaders = headers.filter(header => columnsToKeep.includes(header));
    let filteredData = data.map(row => row.filter((_, index) => columnsToKeep.includes(headers[index])));

    // Generate PPT
    let pptx = new PptxGenJS();
    filteredData.forEach(row => {
        let slide = pptx.addSlide();
        let tableData = [[...filteredHeaders], [...row]];
        slide.addTable(tableData, { x: 0.5, y: 0.5, w: 8.5, h: 5.0 });
    });

    pptx.writeFile({ fileName: "customer_complaints_report.pptx" }).then(() => {
        messageContainer.textContent = 'PPT file has been generated. Check your downloads folder.';
        messageContainer.className = '';
    }).catch(err => {
        messageContainer.textContent = 'Error generating PPT: ' + err.message;
        messageContainer.className = 'error-message';
    });
}
