hiddenUploadBtn = document.getElementById('hiddenUploadBtn');
uploadFile = document.getElementById('uploadFile');
currFile = document.getElementById('currFile');

hiddenUploadBtn.addEventListener('change', function () {
    currFile.textContent = this.files[0].name;
    uploadFile.textContent = 'Change File';

});

