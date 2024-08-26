# pptx 코드 환경 없어서 못하고 있음

$('#uploadButton').click(function (e) {
    e.preventDefault();
    let newFile = $('#newFile')[0].files[0];
    let messageContainer = $('#message');
    let downloadSection = $('#download-section');
    let downloadLink = $('#download-link');
    let progress = $('#progress');

    if (!newFile) {
        messageContainer.text('엑셀 파일을 선택해 주세요.').addClass('error-message');
        return;
    }

    let form = new FormData();
    form.append('file', newFile);

    $.ajax({
        type: 'post',
        url: getWebAppBackendUrl('/upload-to-dss'),
        processData: false,
        contentType: false,
        data: form,
        xhr: function() {
            let xhr = new window.XMLHttpRequest();
            xhr.upload.addEventListener("progress", function(evt) {
                if (evt.lengthComputable) {
                    let pct = parseInt(evt.loaded / evt.total * 100);
                    progress.css("width", "" + pct + "%");
                }
            }, false);
            return xhr;
        },
        success: function (data) {
            messageContainer.text('파일 업로드 및 PPT 생성이 완료되었습니다.').removeClass('error-message');
            downloadLink.attr('href', getWebAppBackendUrl('/download-ppt'));
            downloadSection.show();
        },
        error: function (jqXHR, status, errorThrown) {
            messageContainer.text('업로드 중 오류가 발생했습니다: ' + jqXHR.responseText).addClass('error-message');
        },
        complete: function () {
            progress.css("width", "0%");
        }
    });
});

$.getJSON(getWebAppBackendUrl('/first_api_call'), function(data) {
    console.log('Received data from backend', data)
    const output = $('<pre />').text('Backend reply: ' + JSON.stringify(data));
    $('body').append(output)
});
