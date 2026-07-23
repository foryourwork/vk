<?php
// --- BACKEND PHP LOGIC (AES-128 Encryption & Portal Sync) ---
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_GET['action']) && $_GET['action'] === 'track') {
    header('Content-Type: application/json');
    $input = json_decode(file_get_contents("php://input"), true);
    $appId = trim($input['app_id'] ?? '');

    if (!$appId) {
        echo json_encode(['status' => 'error', 'message' => 'Invalid ID']);
        exit;
    }

    $statusUrl = "https://msme.up.gov.in/Home/Get_ApplicationStatusData";
    $ch = curl_init($statusUrl);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_POST, true);
    curl_setopt($ch, CURLOPT_POSTFIELDS, http_build_query(['username' => $appId]));
    curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
    $statusResponse = curl_exec($ch);
    curl_close($ch);

    $statusData = json_decode($statusResponse, true);

    if (!$statusData || $statusData['status'] === null || $statusData['status'] === "-1") {
        echo json_encode(['status' => 'error', 'message' => 'Record not found']);
        exit;
    }

    echo json_encode(["info" => $statusData]);
    exit;
}
?>

<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bulk MSME Sync Pro | Aman</title>
    <script src="https://www.gstatic.com/firebasejs/9.17.1/firebase-app-compat.js"></script>
    <script src="https://www.gstatic.com/firebasejs/9.17.1/firebase-database-compat.js"></script>
    
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;700;800&display=swap" rel="stylesheet">
    <style>
        :root { --primary: #09575f; --accent: #f35253; --bg: #f5f7fa; }
        body { font-family: 'Plus Jakarta Sans', sans-serif; background: var(--bg); margin: 0; }
        .header { background: var(--primary); color: white; padding: 30px; text-align: center; border-bottom: 5px solid var(--accent); }
        .container { max-width: 1200px; margin: -30px auto 50px; padding: 0 20px; }
        .card { background: white; padding: 30px; border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); border: none; margin-bottom: 20px; }
        .btn-group { display: flex; gap: 10px; justify-content: center; flex-wrap: wrap; }
        .upload-btn { background: var(--accent); color: white; padding: 12px 25px; border-radius: 50px; cursor: pointer; font-weight: bold; transition: 0.3s; border: none; }
        .download-btn { background: #10b981; color: white; padding: 12px 25px; border-radius: 50px; cursor: pointer; font-weight: bold; transition: 0.3s; border: none; display: none; }
        .progress-container { margin-top: 25px; display: none; }
        .progress-bar { height: 12px; background: #e9ecef; border-radius: 10px; overflow: hidden; margin-top: 10px; }
        #fill { width: 0%; height: 100%; background: linear-gradient(90deg, #09575f, #12a4b3); transition: 0.3s; }
        table { width: 100%; border-collapse: collapse; background: white; border-radius: 12px; overflow: hidden; }
        th { background: #09575f; color: white; padding: 15px; text-align: left; font-size: 13px; }
        td { padding: 15px; border-bottom: 1px solid #eee; font-size: 13px; }
        .status-pill { padding: 5px 12px; border-radius: 50px; font-size: 11px; font-weight: bold; background: #e0f2fe; color: #0369a1; border: 1px solid #0369a1; }
    </style>
</head>
<body>

<div class="header">
    <h2 class="fw-bold">AMAN BULK LIVE SYNC v2.0</h2>
    <p>UP MSME Portal (VSSY) Pipeline & Billing Integrated</p>
</div>

<div class="container">
    <div class="card text-center">
        <div class="btn-group">
            <input type="file" id="excelInput" style="display:none" accept=".xlsx">
            <label for="excelInput" class="upload-btn"><i class="fas fa-file-upload me-2"></i>Upload Excel</label>
            <button id="downloadBtn" class="download-btn" onclick="downloadData()"><i class="fas fa-file-download me-2"></i>Download Full Report</button>
        </div>

        <div class="progress-container" id="progBox">
            <div class="d-flex justify-content-between mb-1">
                <span id="p_status" class="fw-bold text-primary small"></span>
                <span id="p_percent" class="fw-bold">0%</span>
            </div>
            <div class="progress-bar"><div id="fill"></div></div>
        </div>
    </div>

    <div style="overflow-x:auto;">
        <table id="dataTable">
            <thead>
                <tr>
                    <th>S.No</th>
                    <th>App ID</th>
                    <th>आवेदक का नाम</th>
                    <th>जिला</th>
                    <th>वर्तमान स्थिति</th>
                    <th>Pipeline</th>
                </tr>
            </thead>
            <tbody id="tableBody">
                <tr><td colspan="6" class="text-center py-5 text-muted">File upload karein...</td></tr>
            </tbody>
        </table>
    </div>
</div>

<script>
// --- FIREBASE CONFIG ---
const firebaseConfig = {
    apiKey: "AIzaSyBfAuBYPAbcxctTQjiRF3TwPE3eNNwYwxk",
    authDomain: "dbform-db4c9.firebaseapp.com",
    databaseURL: "https://dbform-db4c9-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "dbform-db4c9",
    storageBucket: "dbform-db4c9.firebasestorage.app",
    messagingSenderId: "930632470946",
    appId: "1:930632470946:web:c87062b66c96ee6e5762fb"
};
firebase.initializeApp(firebaseConfig);
const db = firebase.database();

let syncData = [];
const activeUser = JSON.parse(sessionStorage.getItem('activeUser')) || { name: "System Bulk" };

document.getElementById('excelInput').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = async (event) => {
        const data = new Uint8Array(event.target.result);
        const wb = XLSX.read(data, {type: 'array'});
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        
        if(rows.length > 0) {
            document.getElementById('progBox').style.display = 'block';
            document.getElementById('downloadBtn').style.display = 'none';
            syncData = []; 
            await processRows(rows);
        }
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

async function processRows(rows) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    let count = 0;

    for (let i = 0; i < rows.length; i++) {
        const appId = rows[i]['App No'] || rows[i]['Application Number'];
        if (!appId) continue;

        document.getElementById('p_status').innerText = `Syncing & Billing: ${appId}`;
        
        try {
            const res = await fetch('?action=track', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ app_id: appId })
            });
            const result = await res.json();

            if (result.info) {
                const info = result.info;
                syncData.push(info); 
                
                // --- PIPELINE & BILLING ENTRY ---
                await db.ref(`absent_sync_pipeline/${info.App_Id}`).set({
                    applicant_name: info.applicant_name,
                    portal_status: info.status_str,
                    sync_time: new Date().toLocaleString(),
                    payment_status: 'Unpaid',
                    manager_name: activeUser.name,
                    mobile: info.mobile_no || ""
                });

                tbody.insertAdjacentHTML('beforeend', `
                    <tr>
                        <td>${i + 1}</td>
                        <td class="fw-bold">${info.App_Id}</td>
                        <td>${info.applicant_name}</td>
                        <td>${info.district_name}</td>
                        <td><span class="status-pill">${info.status_str}</span></td>
                        <td class="text-success fw-bold"><i class="fas fa-check-double"></i> Synced</td>
                    </tr>
                `);
                count++;
            }
        } catch (err) { console.error("Error:", appId); }

        const percent = Math.round(((i + 1) / rows.length) * 100);
        document.getElementById('fill').style.width = percent + '%';
        document.getElementById('p_percent').innerText = percent + '%';
        
        await new Promise(r => setTimeout(r, 600)); 
    }
    
    document.getElementById('p_status').innerText = "Bulk Sync & Billing Done!";
    document.getElementById('downloadBtn').style.display = 'inline-block';
    Swal.fire("Success", `${count} Records Synced and added to Billing Pipeline.`, "success");
}

function downloadData() {
    if(syncData.length === 0) return;
    const ws = XLSX.utils.json_to_sheet(syncData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Portal_Sync_Report");
    XLSX.writeFile(wb, `Bulk_MIS_Report_${Date.now()}.xlsx`);
}
</script>
</body>
</html>