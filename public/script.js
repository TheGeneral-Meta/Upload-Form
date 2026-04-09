let credentials = null;

// Baca file JSON credentials
document.getElementById('jsonFile').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            credentials = JSON.parse(event.target.result);
            const statusDiv = document.getElementById('credsStatus');
            
            if (credentials.client_email && credentials.private_key) {
                statusDiv.innerHTML = `<span class="success">✅ Credentials valid! Email: ${credentials.client_email}</span>`;
                document.getElementById('uploadBtn').disabled = false;
            } else {
                statusDiv.innerHTML = `<span class="error">❌ Invalid credentials file</span>`;
            }
        } catch (error) {
            document.getElementById('credsStatus').innerHTML = `<span class="error">❌ Gagal baca JSON: ${error.message}</span>`;
        }
    };
    reader.readAsText(file);
});

// Toggle input spreadsheet ID
document.getElementById('mode').addEventListener('change', (e) => {
    const existingGroup = document.getElementById('existingGroup');
    existingGroup.style.display = e.target.value === 'existing' ? 'block' : 'none';
});

// Upload file Excel
document.getElementById('uploadBtn').addEventListener('click', async () => {
    const excelFile = document.getElementById('excelFile').files[0];
    if (!excelFile) {
        alert('Pilih file Excel dulu!');
        return;
    }
    
    if (!credentials) {
        alert('Upload credentials JSON dulu!');
        return;
    }
    
    const mode = document.getElementById('mode').value;
    const spreadsheetId = document.getElementById('spreadsheetId').value;
    const shareEmail = document.getElementById('shareEmail').value;
    
    if (mode === 'existing' && !spreadsheetId) {
        alert('Masukkan Spreadsheet ID!');
        return;
    }
    
    // Tambahkan share email ke credentials
    if (shareEmail) {
        credentials.shareWithEmail = shareEmail;
    }
    
    // Baca file Excel sebagai base64
    const reader = new FileReader();
    reader.onload = async (event) => {
        const base64Content = event.target.result.split(',')[1];
        
        document.getElementById('loading').style.display = 'block';
        document.getElementById('result').style.display = 'none';
        document.getElementById('uploadBtn').disabled = true;
        
        try {
            const response = await fetch('/api/upload', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    credentials: credentials,
                    fileContent: base64Content,
                    fileName: excelFile.name,
                    mode: mode,
                    spreadsheetId: spreadsheetId
                })
            });
            
            const result = await response.json();
            
            if (result.success) {
                document.getElementById('resultMsg').innerHTML = result.message;
                document.getElementById('resultLink').href = result.url;
                document.getElementById('result').style.display = 'block';
            } else {
                alert('Error: ' + result.error);
            }
        } catch (error) {
            alert('Gagal upload: ' + error.message);
        } finally {
            document.getElementById('loading').style.display = 'none';
            document.getElementById('uploadBtn').disabled = false;
        }
    };
    
    reader.readAsDataURL(excelFile);
});

// Copy link button
document.getElementById('copyBtn')?.addEventListener('click', () => {
    const link = document.getElementById('resultLink').href;
    navigator.clipboard.writeText(link);
    alert('Link disalin!');
});
