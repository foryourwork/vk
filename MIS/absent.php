<?php
// --- BACKEND PHP LOGIC (Portal Sync) ---
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
    echo json_encode($statusData ? ["info" => $statusData] : ['status' => 'error']);
    exit;
}
?>

<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sync Pro | MIS Manager v3.1</title>
    
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://www.gstatic.com/firebasejs/9.17.1/firebase-app-compat.js"></script>
    <script src="https://www.gstatic.com/firebasejs/9.17.1/firebase-database-compat.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;600;700;800&display=swap" rel="stylesheet">
    
    <style>
        :root { --primary: #be123c; --portal: #09575f; --bg: #f8fafc; --warning: #f59e0b; --success: #059669; }
        body { font-family: 'Plus Jakarta Sans', sans-serif; background: var(--bg); margin: 0; font-size: 13px; }
        .header { background: #fff; padding: 12px 20px; border-bottom: 2px solid #e2e8f0; display: flex; align-items: center; gap: 15px; position: sticky; top: 0; z-index: 100; }
        .container { width: 100%; max-width: 1600px; margin: 0 auto; padding: 15px; box-sizing: border-box; }
        .btn { padding: 8px 14px; border-radius: 8px; border: none; font-weight: 700; cursor: pointer; display: inline-flex; align-items: center; gap: 6px; font-size: 10px; transition: 0.2s; }
        .btn-portal { background: var(--portal); color: white; }
        .btn-track { background: #0ea5e9; color: white; }
        .btn-assign { background: #4f46e5; color: white; }
        .btn-dl { background: #1e293b; color: white; }
        .tab-box { display: flex; gap: 10px; margin-bottom: 15px; }
        .tab { padding: 10px 20px; background: #e2e8f0; border-radius: 30px; cursor: pointer; font-weight: 800; font-size: 11px; }
        .tab.active { background: var(--primary); color: white; }
        .status-pill { padding: 4px 8px; border-radius: 4px; font-size: 9px; font-weight: bold; background: #f1f5f9; border: 1px solid #cbd5e1; }
        .row-special { background: #fffde7 !important; border-left: 5px solid var(--warning); }
        .progress-box { background: white; padding: 15px; border-radius: 12px; margin-bottom: 15px; border: 1px solid #e2e8f0; display: none; }
        table { width: 100%; border-collapse: collapse; background: white; border-radius: 12px; overflow: hidden; }
        th { background: #f8fafc; padding: 12px; text-align: left; font-size: 10px; color: #64748b; border-bottom: 1px solid #e2e8f0; }
        td { padding: 12px; border-bottom: 1px solid #f1f5f9; }
        .reveal-box { display: none; font-weight: 800; color: var(--primary); }
    </style>
</head>
<body>

<div class="header">
    <div style="flex-grow:1">
        <h2 style="margin:0; font-size:16px;">MIS Sync Pro v3.1</h2>
        <div id="uNameDisp" style="font-size:10px; font-weight:700; color:var(--primary);">Verifying...</div>
    </div>
    <div style="display:flex; gap:8px;">
        <button class="btn btn-portal" onclick="autoTrackVisibleList()"><i class="fas fa-magic"></i> Auto Track List</button>
        <button class="btn btn-dl" onclick="promptTargetExcel()"><i class="fas fa-file-download"></i> Target Excel</button>
        <button class="btn btn-dl" onclick="downloadExcel('ABSENT')"><i class="fas fa-file-export"></i> Absent Excel</button>
    </div>
</div>

<div class="container">
    <div id="progContainer" class="progress-box">
        <div style="display:flex; justify-content:space-between; font-size:11px;">
            <span id="progStat">Auto-Tracking...</span>
            <span id="progPer">0%</span>
        </div>
        <div style="height:6px; background:#f1f5f9; margin-top:8px; border-radius:10px; overflow:hidden;">
            <div id="progFill" style="height:100%; background:var(--portal); width:0%;"></div>
        </div>
    </div>

    <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap:15px; margin-bottom:20px;">
        <select id="listSelect" onchange="loadTrainings()" style="padding:12px; border-radius:10px; border:1px solid #ccc;"></select>
        <select id="trainingSelect" onchange="syncLive()" style="padding:12px; border-radius:10px; border:1px solid #ccc;"></select>
    </div>

    <div class="tab-box">
        <div class="tab active" id="tabTarget" onclick="switchTab('TARGET')">TARGET LIST</div>
        <div class="tab" id="tabAbsent" onclick="switchTab('ABSENT')">ABSENT LIST</div>
    </div>

    <div style="overflow-x:auto;">
        <table>
            <thead>
                <tr>
                    <th style="width:250px">Actions</th>
                    <th>App No</th>
                    <th>Candidate Details</th>
                    <th>Aadhar No</th>
                    <th>Mobile</th>
                    <th>Portal Status</th>
                    <th>Calling Info</th>
                    <th>Verifier</th>
                </tr>
            </thead>
            <tbody id="dataTable"></tbody>
        </table>
    </div>
</div>

<script>
// --- CONFIG & INIT ---
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

let masterData = [], specialMap = {}, enrolledApps = new Set(), historyMap = {}, callers = [];
let currentTab = 'TARGET', activeUser = null;

window.onload = async () => {
    const session = sessionStorage.getItem('activeUser');
    if (!session) { window.location.href = 'login.html'; return; }
    activeUser = JSON.parse(session);
    document.getElementById('uNameDisp').innerText = `MIS LEAD: ${activeUser.name.toUpperCase()}`;

    const [metaS, specS, enrollS, histS, empS] = await Promise.all([
        db.ref('OverallListsMetadata').once('value'),
        db.ref('SpecialCandidatesList').once('value'),
        db.ref('BatchEnrollments').once('value'),
        db.ref('CallHistory').once('value'),
        db.ref('EmployeeAuthList').once('value')
    ]);

    specS.forEach(c => { if(c.val().appNo) specialMap[String(c.val().appNo).toUpperCase()] = c.val(); });
    enrollS.forEach(b => { Object.keys(b.val()).forEach(app => enrolledApps.add(String(app).toUpperCase())); });
    empS.forEach(e => { if((e.val().roles || []).includes('Caller')) callers.push({id: e.val().pin, name: e.val().name}); });
    historyMap = histS.val() || {};

    let h = '<option value="">-- Select Master List --</option>';
    metaS.forEach(c => { if(c.val().status !== 'Hide It') h += `<option value="${c.key}">${c.val().listName}</option>`; });
    document.getElementById('listSelect').innerHTML = h;
};

async function loadTrainings() {
    const lid = document.getElementById('listSelect').value;
    db.ref('Trainings').orderByChild('listId').equalTo(lid).once('value', s => {
        let h = '<option value="">-- Select Trade --</option>';
        s.forEach(c => { h += `<option value="${c.val().tradeName}">${c.val().trainingName}</option>`; });
        document.getElementById('trainingSelect').innerHTML = h;
    });
}

function switchTab(t) {
    currentTab = t;
    document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
    document.getElementById(t === 'TARGET' ? 'tabTarget' : 'tabAbsent').classList.add('active');
    renderTable();
}

function syncLive() {
    const lid = document.getElementById('listSelect').value;
    const trade = document.getElementById('trainingSelect').value;
    if(!lid || !trade) return;

    db.ref(`FullDataEntries/${lid}`).orderByChild('tradeName').equalTo(trade).on('value', snap => {
        masterData = [];
        snap.forEach(c => {
            const d = c.val(); d.dbKey = c.key;
            if(!enrolledApps.has(String(d.applicationNumber).toUpperCase())) {
                masterData.push(d);
            }
        });
        renderTable();
    });
}

function renderTable() {
    let h = '';
    const filtered = masterData.filter(d => {
        const isA = (d.isAbsent === true || d.status === "ABSENT");
        return (currentTab === 'ABSENT') ? isA : !isA;
    });

    filtered.forEach(d => {
        const appNo = String(d.applicationNumber).toUpperCase();
        const spec = specialMap[appNo];
        const portalStatus = d.portalStatus || 'Pending Sync';
        const isForwarded = portalStatus === "Forwarded to Training Institute by DIC"; // Button Toggle Criteria
        
        let actBtns = `
            <button class="btn btn-track" title="Track" onclick="trackSingle('${appNo}', '${d.dbKey}')"><i class="fas fa-search"></i></button>
            <button class="btn btn-assign" title="Call" onclick="openAssign('${d.dbKey}')"><i class="fas fa-phone"></i></button>
        `;

        if (currentTab === 'TARGET') {
            if (!isForwarded) {
                actBtns += `<button class="btn" style="background:#ef4444; color:white" onclick="markAbsent('${d.dbKey}')">Absent</button>`; // Show if not forwarded
            } else {
                actBtns += `<span class="status-pill" style="color:var(--success); background:#dcfce7">Forwarded (Locked)</span>`; // Hide button if forwarded
            }
        } else {
            // Absent List Tab
            if (isForwarded) {
                actBtns += `<button class="btn" style="background:var(--success); color:white" onclick="restoreCand('${d.dbKey}')">Restore</button>`; // Active if forwarded
            } else {
                actBtns += `<span class="status-pill" style="color:#64748b">Sync Required</span>`; // Inactive/Hide if not forwarded
            }
        }

        const logs = historyMap[d.dbKey] ? Object.values(historyMap[d.dbKey]).pop() : null;
        const callInfo = logs ? `<span class="status-pill">${logs.status}</span>` : 'No Calls';

        const verifierCell = spec ? `
            <button class="btn" onclick="peekVerifier(this, '${spec.of}')"><i class="fas fa-eye"></i></button>
            <span class="reveal-box"></span>
        ` : '-';

        h += `<tr class="${spec ? 'row-special' : ''}">
            <td><div style="display:flex; gap:4px;">${actBtns}</div></td>
            <td><b>${appNo}</b></td>
            <td><b>${d.applicantName}</b><br><small>${d.fatherName}</small></td>
            <td style="color:#0284c7; font-weight:700;">${d.verifiedAadhar || d.aadhar || '-'}</td>
            <td>${d.mobile || '-'}</td>
            <td><span class="status-pill">${portalStatus}</span></td>
            <td>${callInfo}</td>
            <td>${verifierCell}</td>
        </tr>`;
    });
    document.getElementById('dataTable').innerHTML = h || '<tr><td colspan="8" align="center">No Records Found</td></tr>';
}

function peekVerifier(btn, name) {
    const box = btn.nextElementSibling;
    btn.style.display = 'none';
    box.innerText = name;
    box.style.display = 'inline';
    setTimeout(() => {
        box.style.display = 'none';
        btn.style.display = 'inline-flex';
    }, 5000);
}

// --- EXCEL LOGIC ---
async function promptTargetExcel() {
    const { value: filters } = await Swal.fire({
        title: 'Target Excel Download',
        html: `
            <div style="text-align:left; font-size:12px; margin-bottom:5px;">Aadhar Verification Status:</div>
            <select id="f-status" class="swal2-select" style="width:100%; margin-bottom:15px;">
                <option value="ALL">All Candidates</option>
                <option value="VERIFIED">Verified (Aadhar Matched)</option>
                <option value="PENDING">Pending Verification</option>
            </select>
            <div style="text-align:left; font-size:12px; margin-bottom:5px;">Filter Candidate Type:</div>
            <select id="f-type" class="swal2-select" style="width:100%;">
                <option value="ALL">All Types</option>
                <option value="SPECIAL">Special Only</option>
                <option value="NORMAL">Non-Special Only</option>
            </select>
        `,
        showCancelButton: true,
        preConfirm: () => {
            return {
                status: document.getElementById('f-status').value,
                type: document.getElementById('f-type').value
            }
        }
    });
    if (filters) downloadExcel('TARGET', filters);
}

function downloadExcel(type, filters = null) {
    const lid = document.getElementById('listSelect').value;
    if(!lid) return Swal.fire('Error', 'Master List choose karein', 'error');

    let dataToExport = masterData.filter(d => {
        const isA = (d.isAbsent === true || d.status === "ABSENT");
        return type === 'ABSENT' ? isA : !isA;
    });

    if (type === 'TARGET' && filters) {
        dataToExport = dataToExport.filter(d => {
            const isV = d.isVerified === true;
            const isS = !!specialMap[String(d.applicationNumber).toUpperCase()];
            let matchStatus = (filters.status === 'ALL') || (filters.status === 'VERIFIED' ? isV : !isV);
            let matchType = (filters.type === 'ALL') || (filters.type === 'SPECIAL' ? isS : !isS);
            return matchStatus && matchType;
        });
    }

    const rows = dataToExport.map((d, i) => {
        const spec = specialMap[String(d.applicationNumber).toUpperCase()];
        return {
            "S.No": i + 1,
            "App No": d.applicationNumber,
            "Name": d.applicantName,
            "Father Name": d.fatherName,
            "Aadhar Number": d.verifiedAadhar || d.aadhar || "Pending",
            "Mobile": d.mobile || "-",
            "Portal Status": d.portalStatus || "Not Synced",
            "Aadhar Verified": d.isVerified ? "YES" : "NO",
            "Special Verifier": spec ? spec.of : "-",
            "Last Call": historyMap[d.dbKey] ? Object.values(historyMap[d.dbKey]).pop().status : "-"
        };
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = Object.keys(rows[0] || {}).map(() => ({ wch: 20 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `${type}_Report_${new Date().toLocaleDateString()}.xlsx`);
}

// --- UTILITIES ---
async function autoTrackVisibleList() {
    const lid = document.getElementById('listSelect').value;
    const currentList = masterData.filter(d => (currentTab === 'ABSENT' ? (d.isAbsent || d.status === "ABSENT") : (!d.isAbsent && d.status !== "ABSENT")));
    if (currentList.length === 0) return;

    document.getElementById('progContainer').style.display = 'block';

    for (let i = 0; i < currentList.length; i++) {
        const d = currentList[i];
        document.getElementById('progStat').innerText = `Syncing: ${d.applicationNumber}`;
        try {
            const res = await fetch('?action=track', { method: 'POST', body: JSON.stringify({ app_id: d.applicationNumber }) });
            const resJ = await res.json();
            if (resJ.info) {
                const statusText = resJ.info.status_str;
                await db.ref(`FullDataEntries/${lid}/${d.dbKey}`).update({ portalStatus: statusText, lastSyncTime: new Date().toLocaleString() });
            }
        } catch (e) {}
        
        const per = Math.round(((i + 1) / currentList.length) * 100);
        document.getElementById('progFill').style.width = per + '%';
        document.getElementById('progPer').innerText = per + '%';
        await new Promise(r => setTimeout(r, 600)); 
    }
    document.getElementById('progContainer').style.display = 'none';
}

async function trackSingle(appId, dbKey) {
    const lid = document.getElementById('listSelect').value;
    Swal.fire({ title: 'Tracking...', didOpen: () => Swal.showLoading() });
    try {
        const res = await fetch('?action=track', { method: 'POST', body: JSON.stringify({ app_id: appId }) });
        const resJ = await res.json();
        if (resJ.info) {
            const statusText = resJ.info.status_str;
            await db.ref(`FullDataEntries/${lid}/${dbKey}`).update({ portalStatus: statusText, lastSyncTime: new Date().toLocaleString() });
            Swal.fire('Updated', statusText, 'success');
        }
    } catch (e) { Swal.fire('Error', 'Tracking failed', 'error'); }
}

async function markAbsent(key) {
    const lid = document.getElementById('listSelect').value;
    const cand = masterData.find(c => c.dbKey === key);
    const spec = specialMap[String(cand.applicationNumber).toUpperCase()];
    if (spec) {
        const { isConfirmed } = await Swal.fire({
            title: 'Special Candidate!',
            text: `Ye candidate ${spec.of} dwara verified hai. Kya aap ise ABSENT karna chahte hain?`,
            icon: 'warning', showCancelButton: true
        });
        if (!isConfirmed) return;
    }
    await db.ref(`FullDataEntries/${lid}/${key}`).update({ status: "ABSENT", isAbsent: true, absentMarkedAt: new Date().toLocaleString() });
}

async function restoreCand(key) {
    const lid = document.getElementById('listSelect').value;
    await db.ref(`FullDataEntries/${lid}/${key}`).update({ status: "AVAILABLE", isAbsent: false, absentMarkedAt: null });
    Swal.fire('Restored', 'Moved back to Target List', 'success');
}

async function openAssign(dbKey) {
    const cand = masterData.find(c => c.dbKey === dbKey);
    const { value: callerId } = await Swal.fire({
        title: 'Assign Calling',
        input: 'select',
        inputOptions: callers.reduce((acc, c) => ({...acc, [c.id]: c.name}), {}),
        showCancelButton: true
    });
    if(callerId) {
        await db.ref(`UrgentCalls/${dbKey}`).set({
            candidateId: dbKey, applicantName: cand.applicantName, mobile: cand.mobile,
            assignedTo: callerId, status: 'Pending', date: new Date().toLocaleString()
        });
        Swal.fire('Allotted', 'Caller ko allot kar diya gaya hai', 'success');
    }
}
</script>
</body>
</html>