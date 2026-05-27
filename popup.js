document.addEventListener('DOMContentLoaded', () => {
    const exportBtn = document.getElementById('exportBtn');
    const statusDiv = document.getElementById('status');
    const fileNameInput = document.getElementById('fileNameInput');

    exportBtn.addEventListener('click', async () => {
        statusDiv.style.display = 'block';
        statusDiv.innerText = 'Đang lấy dữ liệu...';

        try {
            let [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

            chrome.scripting.executeScript({
                target: { tabId: tab.id },
                function: extractSwaggerAPI
            }, (injectionResults) => {
                if (chrome.runtime.lastError) {
                    statusDiv.innerText = 'Lỗi: ' + chrome.runtime.lastError.message;
                    return;
                }
                // test4
                // test 
                if (!injectionResults || !injectionResults[0] || !injectionResults[0].result) {
                    statusDiv.innerText = "Không thể trích xuất dữ liệu. Hãy chắc chắn bạn đang ở trang Swagger UI.";
                    return;
                }

                const { apis, pageTitle } = injectionResults[0].result;
                if (!apis || apis.length === 0) {
                    statusDiv.innerText = "Không tìm thấy API nào. Hãy đảm bảo danh sách API đã được render.";
                    return;
                }

                statusDiv.innerText = `Đã tìm thấy ${apis.length} APIs. Đang xuất Excel...`;

                let customName = fileNameInput.value.trim();
                let finalFileName = customName ? customName : (pageTitle ? pageTitle : "Swagger_APIs");
                if (!finalFileName.toLowerCase().endsWith('.xlsx')) {
                    finalFileName += '.xlsx';
                }

                exportToExcel(apis, finalFileName);

                setTimeout(() => {
                    statusDiv.innerText = `Thành công! Đã tải xuống ${apis.length} APIs.`;
                }, 1000);
            });
        } catch (error) {
            statusDiv.innerText = 'Lỗi: ' + error.message;
        }
    });
});

function extractSwaggerAPI() {
    const data = [];

    let pageTitle = document.title || 'Swagger_APIs';
    pageTitle = pageTitle.replace(/[/\\?%*:|"<>]/g, '-').trim();
    const elements = document.querySelectorAll('.opblock');
    if (elements.length === 0) return { apis: [], pageTitle: pageTitle };

    elements.forEach(el => {

        // Method (GET, POST, PUT, DELETE...)
        let methodElem = el.querySelector('.opblock-summary-method');
        let method = methodElem ? methodElem.innerText.trim().toUpperCase() : '';

        // Path (/api/v1/users)
        let pathElem = el.querySelector('.opblock-summary-path');
        let path = '';
        if (pathElem) {
            let dataPath = pathElem.getAttribute('data-path');
            if (dataPath) {
                path = dataPath;
            } else {
                let pathSpan = pathElem.querySelector('a span') || pathElem.querySelector('span');
                path = pathSpan ? pathSpan.innerText.trim() : pathElem.innerText.trim();
            }
        }
        // Loại bỏ ký tự khoảng trắng tàng hình (zero-width spaces) thường gặp trong Swagger
        path = path.replace(/[\u200B-\u200D\uFEFF]/g, '');

        // Description (Summary của API)
        let descElem = el.querySelector('.opblock-summary-description');
        let description = descElem ? descElem.innerText.trim() : '';

        // Name (Lấy tên Tag của section chứa API)
        let section = el.closest('.opblock-tag-section');
        let name = '';
        if (section) {
            let tagElem = section.querySelector('.opblock-tag');
            if (tagElem) {
                name = tagElem.innerText.split('\n')[0].trim();
            }
        }

        // Cấu trúc theo đúng thứ tự: name, path, method, description
        data.push({
            name: name,
            path: path,
            method: method,
            description: description
        });
    });

    return { apis: data, pageTitle: pageTitle };
}

function exportToExcel(data, fileName) {
    if (typeof XLSX === 'undefined') {
        alert("Lỗi tải thư viện SheetJS XLSX.");
        return;
    }

    const worksheet = XLSX.utils.json_to_sheet(data, {
        header: ["name", "path", "method", "description"]
    });

    worksheet['!cols'] = [
        { wch: 25 }, // name
        { wch: 50 }, // path
        { wch: 10 }, // method
        { wch: 60 }  // description
    ];

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Swagger APIs");

    XLSX.writeFile(workbook, fileName);
}
