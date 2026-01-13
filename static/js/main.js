document.getElementById('uploadForm').addEventListener('submit', async (e) => {
    e.preventDefault();

    const formData = new FormData();
    const fileInput = document.getElementById('contractFile');
    formData.append('contract', fileInput.files[0]);

    const analyzeBtn = document.getElementById('analyzeBtn');
    const btnText = document.getElementById('btnText');
    const btnSpinner = document.getElementById('btnSpinner');
    const results = document.getElementById('results');
    const error = document.getElementById('error');

    // Показать загрузку
    btnText.classList.add('d-none');
    btnSpinner.classList.remove('d-none');
    analyzeBtn.disabled = true;
    results.classList.add('d-none');
    error.classList.add('d-none');

    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (data.success) {
            document.getElementById('analysisText').textContent = data.analysis;
            document.getElementById('protocolText').textContent = data.protocol;
            document.getElementById('downloadLink').href = data.download_url;
            results.classList.remove('d-none');
        } else {
            error.textContent = data.error || 'Произошла ошибка при обработке';
            error.classList.remove('d-none');
        }
    } catch (err) {
        error.textContent = 'Ошибка соединения с сервером';
        error.classList.remove('d-none');
    } finally {
        btnText.classList.remove('d-none');
        btnSpinner.classList.add('d-none');
        analyzeBtn.disabled = false;
    }
});
