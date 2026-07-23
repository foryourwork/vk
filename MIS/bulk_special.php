<?php
// --- BACKEND PHP: MSME Portal API Bridge ---
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_GET['action']) && $_GET['action'] === 'portal_sync') {
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
    curl_setopt($ch, CURLOPT_TIMEOUT, 20);
    $statusResponse = curl_exec($ch);
    curl_close($ch);

    echo $statusResponse ?: json_encode(['status' => 'error', 'message' => 'Portal Timeout']);
    exit;
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bulk Merge Pro | Enterprise Edition</title>
    
    <script src="https://www.gstatic.com/firebasejs/9.17.1/firebase-app-compat.js"></script>
    <script src="https://www.gstatic.com/firebasejs/9.17.1/firebase-database-compat.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">

    <style>
        :root { --primary: #4f46e5; --accent: #10b981; --merge: #f59e0b; --dark: #0f172a; }
        body { font-family: 'Plus Jakarta Sans', sans-serif; background: #f8fafc; margin: 0; color: #334155; }
        
        .hero { background: linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%); color: white; padding: 60px 20px 100px; text-align: center; clip-path: ellipse(150% 100% at 50% 0%); }
        .hero h1 { font-size: 2.2rem; font-weight: 800; margin: 0; letter-spacing: -1px; }

        .wrapper { max-width: 1240px; margin: -60px auto 50px; padding: 0 20px; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: white; padding: 20px; border-radius: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); display: flex; align-items: center; gap: 15px; border: 1px solid #e2e8f0; }
        .stat-icon { height: 45px; width: 45px; border-radius: 12px; display: flex; align-items: center; justify-content: center; font-size: 1.2rem; }

        .console-card { background: white; border-radius: 30px; padding: 35px; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.1); border: 1px solid white; }
        
        .upload-box { border: 2px dashed #cbd5e1; border-radius: 24px; padding: 40px; text-align: center; cursor: pointer; transition: 0.3s; background: #f8fafc; border-color: var(--primary); }
        .upload-box:hover { background: #f5f3ff; transform: translateY(-3px); }

        .btn { padding: 14px 28px; border-radius: 16px; border: none; font-weight: 700; cursor: pointer; transition: 0.3s; display: inline-flex; align-items: center; gap: 10px; font-size: 14px; }
        .btn-sample { background: #f1f5f9; color: var(--dark); margin-bottom: 20px; }
        .btn-forward { background: var(--accent); color: white; width: 100%; justify-content: center; font-size: 1.1rem; display: none; margin-top: 25px; }
        .btn-excel { background: #1e293b; color: white; margin-top: 15px; display: none; }

        .table-res { overflow-x: auto; margin-top: 30px; border-radius: 20px; border: 1px solid #e2e8f0; display: none; }
        table { width: 100%; border-collapse: collapse; background: white; }
        th { background: #f8fafc; padding: 18px; text-align: left; font-size: 11px; text-transform: uppercase; color: #64748b; font-weight: 800; }
        td { padding: 16px; border-bottom: 1px solid #f1f5f9; font-size: 13px; }
        
        .badge { padding: 5px 12px; border-radius: 8px; font-size: 10px; font-weight: 800; }
        .status-new { background: #dcfce7; color: #166534; }
        .status-merge { background: #fef3c7; color: #92400e; }
        
        .progress-fill { height: 10px; background: linear-gradient(90deg, #4f46e5, #10b981); transition: 0.4s; border-radius: 10px; width: 0%; }
    </style>
</head>
<body>

<div class="hero">
    <h1><i class="fas fa-sync-alt"></i> Bulk Sync & Merge Engine</h1>
    <p>Portal Status Update & Candidate Data Merging System</p>
</div>

<div class="wrapper">
    <div class="stats-grid">
        <div class="stat-card">
            <div class="stat-icon" style="background: #eef2ff; color: #4338ca;"><i class="fas fa-file-import"></i></div>
            <div><small style="color: #64748b;">Excel Rows</small><div id="stat_rows" style="font-weight: 800; font-size: 1.1rem;">0</div></div>
        </div>
        <div class="stat-card">
            <div class="stat-icon" style="background: #f0fdf4; color: #16a34a;"><i class="fas fa-plus-circle"></i></div>
            <div><small style="color: #64748b;">New Entries</small><div id="stat_new" style="font-weight: 800; font-size: 1.1rem; color: #16a34a;">0</div></div>
        </div>
        <div class="stat-card">
            <div class="stat-icon" style="background: #fffbeb; color: #d97706;"><i class="fas fa-random"></i></div>
            <div><small style="color: #64748b;">Merging (Updates)</small><div id="stat_merge" style="font-weight: 800; font-size: 1.1rem; color: #d97706;">0</div></div>
        </div>
    </div>

    <div class="console-card">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
            <button class="btn btn-sample" onclick="downloadSampleFormat()"><i class="fas fa-download"></i> Get Template</button>
            <div id="userTag" style="font-size: 11px; font-weight: 800; background: #f1f5f9; padding: 8px 15px; border-radius: 12px; color: var(--primary);"></div>
        </div>

        <div class="upload-box" onclick="document.getElementById('xlFile').click()">
            <i class="fas fa-file-excel fa-3x" style="color: var(--primary); margin-bottom: 15px;"></i>
            <h3>Upload Master Excel</h3>
            <p style="font-size: 13px; color: #64748b;">Required: <b>APPLICATION NUMBER, TRADE, VERIFIER</b></p>
        </div>
        <input type="file" id="xlFile" hidden accept=".xlsx, .xls">

        <div id="progBar" style="display:none; margin-top: 25px;">
            <div style="display:flex; justify-content:space-between; font-size:12px; font-weight:800; margin-bottom:8px;">
                <span id="progLabel" style="color:var(--primary);">Processing...</span>
                <span id="progPct">0%</span>
            </div>
            <div style="background:#e2e8f0; height:10px; border-radius:10px;"><div class="progress-fill" id="pFill"></div></div>
        </div>

        <button class="btn btn-excel" id="exportXlBtn" onclick="exportProcessedData()">
            <i class="fas fa-file-export"></i> DOWNLOAD SYNC REPORT
        </button>

        <div class="table-res" id="tableRes">
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>App ID</th>
                        <th>Name</th>
                        <th>Trade/Verifier</th>
                        <th>Portal Status</th>
                        <th>Action</th>
                    </tr>
                </thead>
                <tbody id="syncTbody"></tbody>
            </table>
        </div>

        <button class="btn btn-forward" id="forwardBtn" onclick="executeFinalSync()">
            <i class="fas fa-cloud-upload-alt"></i> SYNC & MERGE ALL DATA
        </button>
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

let finalPayload = [];
let existingDataMap = {}; 
const activeUser = JSON.parse(sessionStorage.getItem('activeUser')) || { name: "Aman_Manager" };
document.getElementById('userTag').innerText = `USER: ${activeUser.name.toUpperCase()}`;

function downloadSampleFormat() {
    const data = [["APPLICATION NUMBER", "TRADE", "VERIFIER"], ["VSSY20250001", "TAILORING", "AMAN SHRIVASTAV"]];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Format");
    XLSX.writeFile(wb, "Bulk_Master_Template.xlsx");
}

document.getElementById('xlFile').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if(!file) return;

    Swal.fire({ title: 'Loading Database...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });

    const snap = await db.ref('SpecialCandidatesList').once('value');
    existingDataMap = {};
    snap.forEach(s => {
        const val = s.val();
        const appNo = val && val.appNo ? String(val.appNo).trim().toUpperCase() : null;
        if(appNo) existingDataMap[appNo] = { key: s.key, data: val };
    });

    const reader = new FileReader();
    reader.onload = async (event) => {
        const wb = XLSX.read(new Uint8Array(event.target.result), {type: 'array'});
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        Swal.close();
        if(rows.length > 0) processBulkSync(rows);
    };
    reader.readAsArrayBuffer(file);
});

async function processBulkSync(rows) {
    document.getElementById('progBar').style.display = 'block';
    document.getElementById('tableRes').style.display = 'block';
    document.getElementById('exportXlBtn').style.display = 'none'; // Reset button
    const tbody = document.getElementById('syncTbody');
    tbody.innerHTML = '';
    finalPayload = [];
    let mergeCount = 0, newCount = 0;

    document.getElementById('stat_rows').innerText = rows.length;

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        let app = (row['APPLICATION NUMBER'] || row['App No'] || "").toString().trim().toUpperCase();
        let trd = row['TRADE'] || "N/A";
        let ver = row['VERIFIER'] || "N/A";

        if(!app) continue;

        const existingRecord = existingDataMap[app];
        const isMerging = !!existingRecord;

        let pct = Math.round(((i + 1) / rows.length) * 100);
        document.getElementById('pFill').style.width = pct + '%';
        document.getElementById('progPct').innerText = pct + '%';
        document.getElementById('progLabel').innerText = `Portal Tracking: ${app}`;

        try {
            const res = await fetch('?action=portal_sync', {
                method: 'POST',
                body: JSON.stringify({ app_id: app })
            });
            const portalData = await res.json();

            if (portalData.applicant_name) {
                let mergedItem = isMerging ? { ...existingRecord.data } : {}; 
                
                // --- UPGRADE: Instant Portal Status & DateTime Logic ---
                const currentTime = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
                
                mergedItem.appNo = app;
                mergedItem.name = portalData.applicant_name;
                mergedItem.status = portalData.status_str; // Instant Portal Status
                mergedItem.trade = trd;
                mergedItem.of = ver;
                mergedItem.father = portalData.father_name || mergedItem.father || "";
                mergedItem.dob = portalData.dob || mergedItem.dob || "";
                mergedItem.mobile = portalData.mobile_no || mergedItem.mobile || "";
                mergedItem.updatedBy = activeUser.name;
                mergedItem.lastSync = currentTime; // Current Sync Date Time
                if(!isMerging) mergedItem.entryAt = currentTime;

                finalPayload.push({
                    key: isMerging ? existingRecord.key : null,
                    data: mergedItem,
                    type: isMerging ? 'Merged/Updated' : 'New Entry' // For Excel Report
                });

                if(isMerging) mergeCount++; else newCount++;

                tbody.insertAdjacentHTML('beforeend', `
                    <tr style="${isMerging ? 'background:#fffbeb;' : ''}">
                        <td>${i+1}</td>
                        <td><b>${app}</b></td>
                        <td>${portalData.applicant_name}</td>
                        <td><small>${trd}<br>${ver}</small></td>
                        <td><span style="color:var(--primary); font-weight:800;">${portalData.status_str}</span><br><small style="font-size:9px;">${currentTime}</small></td>
                        <td><span class="badge ${isMerging ? 'status-merge' : 'status-new'}">${isMerging ? 'Update/Merge' : 'New Entry'}</span></td>
                    </tr>
                `);
            }
        } catch(e) { console.error(e); }

        document.getElementById('stat_new').innerText = newCount;
        document.getElementById('stat_merge').innerText = mergeCount;
        await new Promise(r => setTimeout(r, 300)); // Slightly faster delay
    }

    if(finalPayload.length > 0) {
        document.getElementById('forwardBtn').style.display = 'inline-flex';
        document.getElementById('exportXlBtn').style.display = 'inline-flex';
    }
}

// --- UPGRADE 1: Export Processed Data to Excel ---
function exportProcessedData() {
    if(finalPayload.length === 0) return;

    const exportData = finalPayload.map((item, index) => ({
        "S.No": index + 1,
        "Application ID": item.data.appNo,
        "Candidate Name": item.data.name,
        "Father Name": item.data.father,
        "Trade": item.data.trade,
        "Verifier": item.data.of,
        "Portal Status": item.data.status,
        "Last Sync At": item.data.lastSync,
        "Action Type": item.type
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SyncReport");
    XLSX.writeFile(wb, `Sync_Report_${new Date().getTime()}.xlsx`);
}

async function executeFinalSync() {
    const confirm = await Swal.fire({
        title: 'Execute Sync?',
        text: `You are about to Merge ${document.getElementById('stat_merge').innerText} and Add ${document.getElementById('stat_new').innerText} entries.`,
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'Yes, Sync All'
    });

    if(confirm.isConfirmed) {
        Swal.fire({ title: 'Updating Database...', didOpen: () => Swal.showLoading() });
        const batch = {};
        
        finalPayload.forEach(item => {
            const finalKey = item.key || db.ref('SpecialCandidatesList').push().key;
            batch[`/SpecialCandidatesList/${finalKey}`] = item.data;
        });

        await db.ref().update(batch);
        Swal.fire("Success", "Special List Updated with Portal Status!", "success").then(() => location.reload());
    }
}
</script>
</body>
</html>