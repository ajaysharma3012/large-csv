document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('uploadForm');
    const fileList = document.getElementById('fileList');
    const message = document.getElementById('message');
    const progressBar = document.getElementById('progressBar').firstElementChild;
    const downloadButton = document.getElementById('downloadButton');

    form.addEventListener('submit', function(e) {
        e.preventDefault();
        const masterFile = document.getElementById('masterFile').files[0];
        const secondaryFiles = document.getElementById('secondaryFiles').files;
        const rowsToDelete = parseInt(document.getElementById('rowsToDelete').value) || 0;

        if (!masterFile || secondaryFiles.length === 0) {
            message.textContent = 'Please select all required files.';
            return;
        }

        progressBar.style.width = '0%';
        message.textContent = 'Merging files...';
        downloadButton.style.display = 'none';

        mergeFiles(masterFile, secondaryFiles, rowsToDelete)
            .then(mergedWorkbook => {
                progressBar.style.width = '100%';
                message.textContent = 'Files merged successfully!';
                downloadButton.style.display = 'block';
                downloadButton.onclick = function() {
                    XLSX.writeFile(mergedWorkbook, 'merged.xlsx');
                };
            })
            .catch(error => {
                console.error('Error:', error);
                message.textContent = 'An error occurred while merging files.';
                progressBar.style.width = '0%';
            });
    });

    document.getElementById('secondaryFiles').addEventListener('change', function(e) {
        fileList.innerHTML = '';
        for (let i = 0; i < this.files.length; i++) {
            fileList.innerHTML += `<p>${this.files[i].name}</p>`;
        }
    });

    async function mergeFiles(masterFile, secondaryFiles, rowsToDelete) {
        const masterData = await readFile(masterFile);
        const secondaryDataPromises = Array.from(secondaryFiles).map(file => readFile(file, rowsToDelete));
        const secondaryDataArrays = await Promise.all(secondaryDataPromises);

        const mergedData = masterData.concat(...secondaryDataArrays);
        const mergedWorkbook = XLSX.utils.book_new();
        const mergedSheet = XLSX.utils.aoa_to_sheet(mergedData);
        XLSX.utils.book_append_sheet(mergedWorkbook, mergedSheet, "Merged Data");

        return mergedWorkbook;
    }

    function readFile(file, rowsToDelete = 0) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
                    jsonData = jsonData.slice(rowsToDelete);
                    resolve(jsonData);
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }
});