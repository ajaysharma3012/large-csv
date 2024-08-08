document.addEventListener('DOMContentLoaded', function () {
    const form = document.getElementById('uploadForm');
    const fileList = document.getElementById('fileList');
    const message = document.getElementById('message');
    const progressBar = document.getElementById('progressBar').firstElementChild;
    const downloadButton = document.getElementById('downloadButton');
  
    form.addEventListener('submit', async function (e) {
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
  
      try {
        const mergedBlob = await callNetlifyFunction(masterFile, secondaryFiles, rowsToDelete);
        progressBar.style.width = '100%';
        message.textContent = 'Files merged successfully!';
        downloadButton.style.display = 'block';
        downloadButton.onclick = function () {
          saveAs(mergedBlob, 'merged.xlsx');
        };
      } catch (error) {
        console.error('Error:', error);
        message.textContent = 'An error occurred while merging files.';
        progressBar.style.width = '0%';
      }
    });
  
    document.getElementById('secondaryFiles').addEventListener('change', function (e) {
      fileList.innerHTML = '';
      for (let i = 0; i < this.files.length; i++) {
        fileList.innerHTML += `<p>${this.files[i].name}</p>`;
      }
    });
  
    async function callNetlifyFunction(masterFile, secondaryFiles, rowsToDelete) {
      const formData = new FormData();
      formData.append('masterFile', masterFile);
      for (let i = 0; i < secondaryFiles.length; i++) {
        formData.append('secondaryFiles', secondaryFiles[i]);
      }
      formData.append('rowsToDelete', rowsToDelete);
  
      const response = await fetch('/.netlify/functions/merge-files', {
        method: 'POST',
        body: JSON.stringify({ masterFile, secondaryFiles, rowsToDelete }),
      });
  
      if (!response.ok) {
        throw new Error('Network response was not ok');
      }
  
      return await response.blob();
    }
  
    function saveAs(blob, filename) {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.style.display = 'none';
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    }
  });