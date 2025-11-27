<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="matchActualPO.aspx.vb" Inherits="BMS.matchActualPO" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - Match Actual PO</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="style/theme.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <style>
        /* Override styles specific to this page if needed */
        /* Ensure highlight classes are working even if theme.css fails */
        .highlight-matched {
            background-color: #d4edda !important; /* Green for Matched */
        }
        .highlight-matching {
            background-color: #fff3cd !important; /* Yellow for Matching */
        }

        /* [BMS Gem] New Header Styles */
        .table thead th {
            vertical-align: middle !important;
            text-align: center;
            border-color: #fff !important; /* เส้นคั่นสีขาว */
        }

        /* 1. Common Keys (Grey/Dark) */
        .th-key {
            background-color: #5a6268 !important; 
            color: white !important;
        }

        /* 2. Draft PO Group (Teal/Light Blue) */
        .th-draft {
            background-color: #17a2b8 !important; 
            color: white !important;
        }

        /* 3. Actual PO Group (Primary Blue) */
        .th-actual {
            background-color: var(--primary-blue) !important; 
            color: white !important;
        }
    </style>
</head>
<body>
     <!-- (เพิ่ม) Loading Overlay -->
    <div class="loading-overlay" id="loadingOverlay">
        <div class="loading-content">
            <div class="loading-spinner"></div>
            <p class="loading-text" id="loadingText">Processing...</p>
        </div>
    </div>
    <!-- Sidebar Overlay -->
    <div class="sidebar-overlay" id="sidebarOverlay" onclick="toggleSidebar()"></div>

    <!-- Sidebar -->
    <div class="sidebar" id="sidebar">
        <div class="sidebar-header">
            <h3><a class="text-decoration-none text-white" href="dashboard.aspx" ><i class="bi bi-building"></i> KBMS</a></h3>
            <button class="close-sidebar" onclick="toggleSidebar()">
                <i class="bi bi-x-lg"></i>
            </button>
        </div>
        <ul class="sidebar-menu">
            <li class="menu-item">
                <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'otbPlan')">
                    <i class="bi bi-clipboard-data"></i>
                    <span>OTB Plan / Revise</span>
                    <i class="bi bi-chevron-down"></i>
                </a>
                <ul class="submenu" id="otbPlan">
                    <li><a href="draftOTB.aspx" class="menu-link">Draft OTB Plan</a></li>
                    <li><a href="approvedOTB.aspx" class="menu-link">Approved OTB Plan</a></li>
                </ul>
            </li>
            <li class="menu-item">
                <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'otbSwitching')">
                    <i class="bi bi-arrow-left-right"></i>
                    <span>OTB Switching</span>
                    <i class="bi bi-chevron-down"></i>
                </a>
                <ul class="submenu" id="otbSwitching">
                    <li><a href="createOTBswitching.aspx" class="menu-link">Create OTB Switching</a></li>
                    <li><a href="transactionOTBSwitching.aspx" class="menu-link">Switching Transaction</a></li>
                </ul>
            </li>
            <li class="menu-item">
                <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'po')">
                    <i class="bi bi-file-earmark-text"></i>
                    <span>PO</span>
                    <i class="bi bi-chevron-down"></i>
                </a>
                <ul class="submenu" id="po">
                    <li><a href="createDraftPO.aspx" class="menu-link">Create Draft PO</a></li>
                    <li><a href="draftPO.aspx" class="menu-link">Draft PO</a></li>
                    <li><a href="matchActualPO.aspx" class="menu-link active">Match Actual PO</a></li>
                    <li><a href="actualPO.aspx" class="menu-link">Actual PO</a></li>
                </ul>
            </li>
            <li class="menu-item">
                <a href="otbRemaining.aspx" class="menu-link">
                    <i class="bi bi-bar-chart-line"></i>
                    <span>OTB Remaining</span>
                </a>
            </li>
            <li class="menu-item">
                <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'master')">
                    <i class="bi bi-database"></i>
                    <span>Master File</span>
                    <i class="bi bi-chevron-down"></i>
                </a>
                <ul class="submenu" id="master">
                    <li><a href="master_vendor.aspx" class="menu-link">Master Vendor</a></li>
                    <li><a href="master_brand.aspx" class="menu-link">Master Brand</a></li>
                    <li><a href="master_category.aspx" class="menu-link">Master Category</a></li>
                </ul>
            </li>
        </ul>
    </div>

    <!-- Main Wrapper -->
    <div class="main-wrapper">
        <!-- Top Navigation -->
        <div class="top-navbar">
            <div class="d-flex align-items-center gap-3">
                <button class="menu-toggle" onclick="toggleSidebar()">
                    <i class="bi bi-list"></i>
                </button>
                <h1 class="page-title" id="pageTitle">KBMS - Match Actual PO</h1>
            </div>
            <div class="user-info">
                <span class="d-none d-md-inline">Welcome, <%= HttpUtility.JavaScriptStringEncode(Session("user").ToString()) %></span>
                <div class="user-avatar">
                    <i class="bi bi-person-circle"></i>
                </div>
            </div>
        </div>

        <!-- Content Area -->
        <div class="content-area">
            <!-- Page Header -->
            <div class="page-header">
                Matching - Actual PO VS Draft PO
            </div>
            <!-- Submit Section -->
            <div class="submit-section">
                <button type="button" id="btnViewNoSync" class="btn-submit me-3" style="background-color: #17a2b8;">
                    <i class="bi bi-table"></i> View Data
                </button>

                <button type="button" id="btnSyncSAP" class="btn-submit me-3" style="background-color: var(--primary-blue);">
                    <i class="bi bi-arrow-repeat"></i> Sync SAP & View
                </button>

                <button type="button" id="btnSubmit" class="btn-submit" disabled>
                    <i class="bi bi-check-circle"></i> Submit
                </button>
            </div>
            <!-- Review Section -->
            
            <!-- Data Table -->
            <div class="table-container">
                <div class="review-section">
                    Review
                </div>
                <div class="table-responsive">
                    <table class="table table-hover mb-0">
                        <thead>
                            <tr>
                                <th class="th-key">Select</th>
                                <th class="th-key">Year</th>
                                <th class="th-key">Month</th>
                                <th class="th-key">Cate</th>
                                <th class="th-key">Company</th>
                                <th class="th-key">Segment</th>
                                <th class="th-key">Brand</th>
                                <th class="th-key">Vendor</th>

                                <th class="th-draft">Draft PO Date</th>
                                <th class="th-draft">Draft PO No.</th>
                                <th class="th-draft">Draft PO Amount (THB)</th>
                                <th class="th-draft">Draft PO Amount (CCY)</th>

                                <th class="th-actual">Actual Amount (CCY)</th>
                                <th class="th-actual">Actual PO Amount (THB)</th>
                                <th class="th-actual">CCY</th>
                                <th class="th-actual">Actual Ex. Rate</th>
                                <th class="th-actual">Actual PO Date</th>
                                <th class="th-actual">Actual PO no.</th>
                            </tr>
                        </thead>
                         <tbody id="matchTableBody">
                            <tr>
                                <td colspan="18" class="text-center p-4 text-muted">
                                    Please click 'Sync SAP' to load data.
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="manualMatchModal" tabindex="-1" aria-labelledby="manualMatchModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-dialog-centered">
            <div class="modal-content">
                <div class="modal-header bg-primary text-white">
                    <h5 class="modal-title" id="manualMatchModalLabel">
                        <i class="bi bi-link-45deg"></i> Manual Match Draft PO
                    </h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <input type="hidden" id="hdnMatchActualPOID"> 
                    <input type="hidden" id="hdnMatchActualPONo">
                    <input type="hidden" id="hdnMatchRowIndex">

                    
                    <div class="mb-3">
                        <label for="txtManualDraftPONo" class="form-label">Actual PO No.</label>
                        <input type="text" class="form-control" id="lblMatchActualPONo" readonly style="background-color: #e9ecef;">
                    </div>
                    <div class="mb-3">
                        <label for="txtManualDraftPONo" class="form-label">Enter Draft PO No. <span class="text-danger">*</span></label>
                        <div class="input-group">
                            <span class="input-group-text"><i class="bi bi-file-earmark-text"></i></span>
                            <input type="text" class="form-control" id="txtManualDraftPONo" placeholder="e.g. DPO-2025-XXXX" autocomplete="off">
                        </div>
                        <div class="form-text text-muted">System will validate if this Draft PO exists.</div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                    <button type="button" class="btn btn-success" onclick="saveManualMatch()">
                        <i class="bi bi-save"></i> Save & Match
                    </button>
                </div>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // (MODIFIED: Cache all buttons)
        let btnSyncSAP = document.getElementById("btnSyncSAP");
        let btnSubmit = document.getElementById("btnSubmit");
        let btnViewNoSync = document.getElementById("btnViewNoSync");
        let tableBody = document.getElementById("matchTableBody");
        let loadingOverlay = document.getElementById("loadingOverlay");
        var manualMatchModal;


        // Toggle Sidebar
        function toggleSidebar() {
            const sidebar = document.getElementById('sidebar');
            const overlay = document.getElementById('sidebarOverlay');
            sidebar.classList.toggle('active');
            overlay.classList.toggle('active');
        }

        // Toggle Submenu
        function toggleSubmenu(event, submenuId) {
            event.preventDefault();
            event.stopPropagation();

            const submenu = document.getElementById(submenuId);
            const menuLink = event.currentTarget;

            submenu.classList.toggle('show');
            menuLink.classList.toggle('expanded');
        }

        // Load Page
        function loadPage(event, pageName) {
            event.preventDefault();

            document.querySelectorAll('.submenu .menu-link').forEach(link => {
                link.classList.remove('active');
            });

            event.currentTarget.classList.add('active');
            document.getElementById('pageTitle').textContent = pageName;

            if (window.innerWidth <= 768) {
                toggleSidebar();
            }
        }

        // (เพิ่ม) Function สำหรับแสดง/ซ่อน Loading
        function showLoading(show, text = 'Processing...') {
            document.getElementById('loadingText').textContent = text;
            if (show) {
                loadingOverlay.classList.add('active');
            } else {
                loadingOverlay.classList.remove('active');
            }
        }

        // (เพิ่ม) Function สำหรับแปลงเลขเป็น format
        function formatNumber(num) {
            if (num === null || num === undefined) return '';
            return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }

        // (เพิ่ม) Function สำหรับแปลงเลขเดือนเป็นชื่อ
        function getMonthName(month) {
            const names = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            const monthInt = parseInt(month, 10);
            return names[monthInt] || month;
        }

        // แก้ไขฟังก์ชัน buildTable ใน matchActualPO.aspx

        function buildTable(data) {
            tableBody.innerHTML = ''; // Clear table

            if (!data || data.length === 0) {
                tableBody.innerHTML = '<tr><td colspan="18" class="text-center p-4 text-muted">No data found.</td></tr>';
                btnSubmit.disabled = true;
                return;
            }

            // Helper สำหรับดึงค่าอย่างปลอดภัย (รองรับทั้งตัวเล็ก/ตัวใหญ่)
            const getVal = (obj, prop) => obj ? (obj[prop] || obj[prop.toLowerCase()] || '') : '';
            const getNum = (obj, prop) => obj ? (obj[prop] || obj[prop.toLowerCase()] || 0) : 0;

            let html = '';
            data.forEach((item, index) => {
                // รองรับทั้ง Key/key, Actual/actual, Draft/draft
                const key = item.Key || item.key;
                const draft = item.Draft || item.draft;
                const actual = item.Actual || item.actual;
                const matchStatus = item.MatchStatus || item.matchStatus;
                const actualPONo = getVal(actual, 'ActualPONo');
                const actualPOID = getVal(actual, 'ActualPO_ID');

                // [BMS Gem Logic Fix]: กำหนดสีตามสถานะ และ ตรวจสอบว่า Lock หรือไม่
                let rowClass = '';
                const isMatched = matchStatus.toLowerCase() === 'matched';
                const isMatching = matchStatus.toLowerCase() === 'matching';

                if (isMatched) {
                    rowClass = 'table-success';
                } else if (isMatching) {
                    rowClass = 'table-warning'; // สีเหลืองสำหรับรายการรอ Submit
                }

                const rowId = `row-${actualPOID}`;

                html += `<tr class="${rowClass}" id="${rowId}">`;

                // Checkbox logic:
                // - ถ้า Matched -> Checked + Disabled (แก้ไขไม่ได้)
                // - ถ้า Matching -> Checked + Enabled (รอ Submit, ผู้ใช้เลือกไม่ส่งได้)
                // - อื่นๆ -> Unchecked + Enabled
                let isChecked = (isMatched || isMatching) ? 'checked' : '';
                let isDisabled = isMatched ? 'disabled' : ''; // [BMS Gem Fix] ปิดการใช้งานถ้า Matched แล้ว

                // --- [BMS Gem Logic] Check Amount CCY Difference ---
                let highlightCCYStyle = '';
                const draftAmtCCY = getNum(draft, 'DraftAmountCCY');
                const actualAmtCCY = getNum(actual, 'ActualAmountCCY');

                // ถ้ามีคู่ Match และยอดเงิน CCY ต่างกัน (มากกว่า 0.01 เพื่อกันเรื่องทศนิยม)
                if (draft && Math.abs(draftAmtCCY - actualAmtCCY) > 0.01) {
                    highlightCCYStyle = 'background-color: #fff3cd; color: #856404; font-weight: bold;'; // สีเหลือง (Warning Theme)
                }

                html += `<td class="text-center" style="white-space: nowrap;">
                            <div class="d-flex align-items-center gap-2 justify-content-center">
                                <input type="checkbox" 
                                       class="form-check-input match-checkbox" 
                                       ${isChecked}
                                       ${isDisabled}
                                       data-draft-pos="${draft ? (draft.DraftPONo || draft.draftPONo) : ''}"
                                       data-actual-po="${actualPONo}"
                                       data-status="${matchStatus}"
                                >
                                <button type="button" class="btn btn-sm btn-outline-secondary" 
                                        onclick="openManualMatchModal('${actualPOID}', '${actualPONo}', ${index})"
                                        title="Change Draft PO No."
                                        ${isDisabled}>
                                    <i class="bi bi-pencil-square"></i>
                                </button>
                            </div>
                         </td>`;

                // --- Group Keys (จาก Actual Summary) ---
                html += `<td>${getVal(key, 'Year')}</td>`;
                html += `<td>${getMonthName(getVal(key, 'Month'))}</td>`;
                html += `<td>${getVal(key, 'Category')}</td>`;
                html += `<td>${getVal(key, 'Company')}</td>`;
                html += `<td>${getVal(key, 'Segment')}</td>`;
                html += `<td>${getVal(key, 'Brand')}</td>`;
                html += `<td>${getVal(key, 'Vendor')}</td>`;

                // --- Draft PO Data (ใช้ ID อ้างอิง Element) ---
                html += `<td id="td-date-${actualPOID}">${draft ? getVal(draft, 'DraftPODate') : '-'}</td>`;
                html += `<td id="td-no-${actualPOID}">${draft ? getVal(draft, 'DraftPONo') : '-'}</td>`;
                html += `<td id="td-thb-${actualPOID}" class="text-end">${draft ? formatNumber(getNum(draft, 'DraftAmountTHB')) : '-'}</td>`;

                html += `<td id="td-ccy-${actualPOID}" class="text-end" style="${highlightCCYStyle}">${draft ? formatNumber(draftAmtCCY) : '-'}</td>`;

                // --- Actual PO Data (ต้องมีค่าเสมอ) ---
                html += `<td class="text-end" style="${highlightCCYStyle}">${formatNumber(actualAmtCCY)}</td>`;

                html += `<td class="text-end"><strong>${formatNumber(getNum(actual, 'ActualAmountTHB'))}</strong></td>`;
                html += `<td>${getVal(actual, 'ActualCCY')}</td>`;
                html += `<td class="text-end">${formatNumber(getNum(actual, 'ActualExRate'), 4)}</td>`; // 4 decimal places for Rate
                html += `<td>${getVal(actual, 'ActualPODate')}</td>`;
                html += `<td>${getVal(actual, 'ActualPONo')}</td>`;

                html += '</tr>';
            });

            tableBody.innerHTML = html;
            btnSubmit.disabled = false;
        }

        // Helper Format Number (ปรับปรุงให้รองรับทศนิยมตามที่ส่งมา)
        function formatNumber(num, decimals = 2) {
            if (num === null || num === undefined || num === '') return '0.00';
            return parseFloat(num).toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
        }

        // (MODIFIED: Helper for encoding attributes)
        const HttpUtility = {
            HtmlAttributeEncode: function (text) {
                if (!text) return '';
                return text.toString()
                    .replace(/"/g, '&quot;')
                    .replace(/'/g, '&#39;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;');
            }
        };

        // (MODIFIED: Refactored sync logic into its own function)
        function syncAndLoadData() {
            showLoading(true, 'Syncing with SAP...');
            btnSubmit.disabled = true; // Disable submit during sync
            btnSyncSAP.disabled = true;

            $.ajax({
                url: 'Handler/POMatchingHandler.ashx?action=get_only',
                type: 'POST',
                processData: false,
                contentType: false,
                dataType: 'json',
                success: function (response) {
                    showLoading(false);
                    btnSyncSAP.disabled = false;
                    if (response.success) {
                        console.log(response.data);
                        buildTable(response.data); // This will re-enable submit if data exists
                        // (MODIFIED: Show sync stats)
                        alert(response.syncStats || "Sync completed.");
                    } else {
                        alert('Error: ' + response.message);
                        tableBody.innerHTML = `<tr><td colspan="18" class="text-center p-4 text-danger">${response.message}</td></tr>`;
                    }
                },
                error: function (xhr, status, error) {
                    showLoading(false);
                    btnSyncSAP.disabled = false;
                    console.error(error);
                    alert('Fatal error connecting to handler: ' + xhr.responseText);
                    tableBody.innerHTML = `<tr><td colspan="18" class="text-center p-4 text-danger">Fatal error: ${xhr.responseText}</td></tr>`;
                }
            });
        }

        // (MODIFIED: New function to handle submit)
        function handleSubmit() {
            let selectedMatches = [];

            // Find all *checked* checkboxes
            document.querySelectorAll('.match-checkbox:checked').forEach(cb => {
                selectedMatches.push({
                    DraftPOs: cb.dataset.draftPos, // "PO-001"
                    ActualPO: cb.dataset.actualPo  // "SAP-12345"
                });
            });

            if (selectedMatches.length === 0) {
                alert("Please select at least one matched row to submit.");
                return;
            }

            if (!confirm(`Are you sure you want to submit ${selectedMatches.length} matched group(s)? This will update the Draft POs.`)) {
                return;
            }

            showLoading(true, `Submitting ${selectedMatches.length} match(es)...`);
            btnSubmit.disabled = true;
            btnSyncSAP.disabled = true;

            $.ajax({
                url: 'Handler/POMatchingHandler.ashx?action=submitmatches',
                type: 'POST',
                data: {
                    matches: JSON.stringify(selectedMatches)
                },
                dataType: 'json',
                success: function (response) {
                    showLoading(false);
                    if (response.success) {
                        alert(response.message);
                        // Refresh the data table
                        syncAndLoadData();
                    } else {
                        alert('Error submitting: ' + response.message);
                        btnSubmit.disabled = false; // Re-enable on error
                        btnSyncSAP.disabled = false;
                    }
                },
                error: function (xhr, status, error) {
                    showLoading(false);
                    btnSubmit.disabled = false; // Re-enable on error
                    btnSyncSAP.disabled = false;
                    alert('Fatal error during submit: ' + xhr.responseText);
                }
            });
        }

        function loadMatchData(actionType) {
            let loadingMsg = actionType === 'sync_and_get' ? 'Syncing SAP & Loading...' : 'Loading Data...';
            showLoading(true, loadingMsg);

            // ปิดปุ่มกันกดซ้ำ
            if (btnSyncSAP) btnSyncSAP.disabled = true;
            if (btnViewNoSync) btnViewNoSync.disabled = true;
            if (btnSubmit) btnSubmit.disabled = true;

            // *** KEY POINT: ไม่ต้องสร้าง FormData จาก Filter แล้ว ***

            $.ajax({
                url: 'Handler/POMatchingHandler.ashx?action=' + actionType, // ส่ง Action ตามปุ่มที่กด
                type: 'POST',
                // data: formData,  <-- เอาออกเลย เพราะ GetPOData ไม่รับค่าแล้ว
                processData: false,
                contentType: false,
                dataType: 'json',
                success: function (response) {
                    showLoading(false);
                    if (btnSyncSAP) btnSyncSAP.disabled = false;
                    if (btnViewNoSync) btnViewNoSync.disabled = false;

                    if (response.success) {
                        buildTable(response.data);
                    } else {
                        alert('Error: ' + response.message);
                        tableBody.innerHTML = `<tr><td colspan="15" class="text-center p-4 text-danger">${response.message}</td></tr>`;
                    }
                },
                error: function (xhr, status, error) {
                    showLoading(false);
                    if (btnSyncSAP) btnSyncSAP.disabled = false;
                    if (btnViewNoSync) btnViewNoSync.disabled = false;
                    console.error(error);
                    alert('Fatal error: ' + xhr.responseText);
                }
            });
        }
        function openManualMatchModal(actualPOID, actualPONo, rowIndex) {
            document.getElementById('hdnMatchActualPOID').value = actualPOID; // Store ID
            document.getElementById('hdnMatchActualPONo').value = actualPONo;
            document.getElementById('lblMatchActualPONo').value = actualPONo;
            document.getElementById('hdnMatchRowIndex').value = rowIndex;
            document.getElementById('txtManualDraftPONo').value = '';
            manualMatchModal.show();

            setTimeout(() => document.getElementById('txtManualDraftPONo').focus(), 500);
        }

        // บันทึกข้อมูล
        function saveManualMatch() {
            var actualPOID = document.getElementById('hdnMatchActualPOID').value;
            var actualPONo = document.getElementById('hdnMatchActualPONo').value;
            var draftPONo = document.getElementById('txtManualDraftPONo').value.trim();

            if (!draftPONo) {
                alert("Please enter Draft PO No.");
                return;
            }

            showLoading(true, "Validating & Matching...");

            var formData = new FormData();
            formData.append('actualPO_ID', actualPOID); // ส่ง ID ไป Backend
            formData.append('actualPONo', actualPONo);
            formData.append('draftPONo', draftPONo);

            $.ajax({
                url: 'Handler/POMatchingHandler.ashx?action=manualMatch',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                dataType: 'json',
                success: function (response) {
                    showLoading(false);
                    if (response.success) {
                        manualMatchModal.hide();
                        alert(response.message);

                        // --- Update UI by ID ---
                        updateRowUI(actualPOID, response.data);

                    } else {
                        alert(response.message);
                    }
                },
                error: function (xhr, status, error) {
                    showLoading(false);
                    alert('Error: ' + (xhr.responseJSON ? xhr.responseJSON.message : error));
                }
            });
        }

        // ฟังก์ชันอัปเดตแถว (ใช้ ID)
        function updateRowUI(actualPOID, newData) {
            var row = document.getElementById('row-' + actualPOID);
            if (row) {
                document.getElementById('td-date-' + actualPOID).textContent = newData.DraftPODate;
                document.getElementById('td-no-' + actualPOID).textContent = newData.DraftPONo;
                document.getElementById('td-thb-' + actualPOID).textContent = formatNumber(newData.DraftAmountTHB);
                document.getElementById('td-ccy-' + actualPOID).textContent = formatNumber(newData.DraftAmountCCY);

                row.classList.add('highlight-matched');

                var checkbox = row.querySelector('.match-checkbox');
                if (checkbox) {
                    checkbox.dataset.draftPos = newData.DraftPONo;
                    checkbox.checked = true;
                }
            }
        }

        let initial = function () {
            // (MODIFIED: Call the refactored sync function)
            manualMatchModal = new bootstrap.Modal(document.getElementById('manualMatchModal'));

            btnSyncSAP.addEventListener('click', function () { loadMatchData('sync_and_get'); });

            // (MODIFIED: Add click listener for submit)
            btnViewNoSync.addEventListener('click', function () { loadMatchData('get_only'); });

            btnSubmit.addEventListener('click', handleSubmit);
        }
        // Close sidebar when clicking outside
        document.addEventListener('click', function (event) {
            const sidebar = document.getElementById('sidebar');
            const menuToggle = document.querySelector('.menu-toggle');

            if (!sidebar.contains(event.target) && !menuToggle.contains(event.target)) {
                if (sidebar.classList.contains('active')) {
                    toggleSidebar();
                }
            }
        });

        // Initialize
        document.addEventListener('DOMContentLoaded', initial);
    </script>
</body>
</html>