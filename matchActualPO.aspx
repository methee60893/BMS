<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="matchActualPO.aspx.vb" Inherits="BMS.matchActualPO" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BMS - Match Actual PO</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <style>
        :root {
            --primary-blue: #0B56A4;
            --sidebar-bg: #2c3e50;
            --sidebar-hover: #34495e;
            --orange-header: #D2691E;
            --light-blue-bg: #E6F2FF;
            --table-header: #4A90E2;
            --green-btn: #28a745;
            --pink-highlight: #FFB6C1;
            /* (เพิ่ม) สีสำหรับแถวที่ Match */
            --green-highlight: #d4edda; 
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f5f5f5;
            overflow-x: hidden;
        }

        /* ( ... CSS อื่นๆ เหมือนเดิม ... ) */
        
        /* (เพิ่ม) สีสำหรับแถวที่ Match */
        .highlight-matched {
             background-color: var(--green-highlight);
        }

        /* Sidebar Styles */
        .sidebar {
            position: fixed;
            top: 0;
            left: -280px;
            height: 100vh;
            width: 280px;
            background: var(--sidebar-bg);
            transition: left 0.3s ease;
            z-index: 2000;
            overflow-y: auto;
            box-shadow: 2px 0 10px rgba(0,0,0,0.3);
        }

        .sidebar.active {
            left: 0;
        }

        .sidebar-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            z-index: 1500;
            display: none;
        }

        .sidebar-overlay.active {
            display: block;
        }

        .sidebar-header {
            padding: 25px 20px;
            background: #1a252f;
            color: white;
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-bottom: 2px solid #34495e;
        }

        .sidebar-header h3 {
            margin: 0;
            font-size: 1.8rem;
            font-weight: 700;
        }

        .close-sidebar {
            background: transparent;
            border: none;
            color: white;
            font-size: 1.8rem;
            cursor: pointer;
            padding: 0;
            width: 35px;
            height: 35px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 4px;
            transition: background 0.3s;
        }

        .close-sidebar:hover {
            background: rgba(255,255,255,0.1);
        }

        .sidebar-menu {
            list-style: none;
            padding: 15px 0;
            margin: 0;
        }

        .menu-item {
            margin: 0;
        }

        .menu-link {
            display: flex;
            align-items: center;
            padding: 14px 20px;
            color: #ecf0f1;
            text-decoration: none;
            transition: all 0.3s ease;
            font-size: 0.95rem;
        }

        .menu-link:hover {
            background: var(--sidebar-hover);
            color: white;
            padding-left: 25px;
        }

        .menu-link.active {
            background: #FF99CC;
            color: white;
            border-left: 4px solid #fff;
        }

        .menu-link i {
            font-size: 1.2rem;
            min-width: 35px;
        }

        .menu-link .bi-chevron-down {
            margin-left: auto;
            font-size: 1rem;
            transition: transform 0.3s;
        }

        .menu-link.expanded .bi-chevron-down {
            transform: rotate(180deg);
        }

        .submenu {
            list-style: none;
            padding-left: 0;
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.4s ease;
            background: rgba(0,0,0,0.2);
        }

        .submenu.show {
            max-height: 600px;
        }

        .submenu .menu-link {
            padding-left: 55px;
            font-size: 0.9rem;
            padding-top: 12px;
            padding-bottom: 12px;
        }

        .submenu .menu-link:hover {
            padding-left: 60px;
        }

        /* Main Wrapper */
        .main-wrapper {
            width: 100%;
            min-height: 100vh;
        }

        /* Top Navigation */
        .top-navbar {
            background: white;
            padding: 15px 30px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: sticky;
            top: 0;
            z-index: 1000;
        }

        .menu-toggle {
            background: var(--primary-blue);
            border: none;
            color: white;
            font-size: 1.5rem;
            cursor: pointer;
            padding: 8px 15px;
            border-radius: 6px;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .menu-toggle:hover {
            background: #094580;
            transform: scale(1.05);
        }

        .page-title {
            color: #2c3e50;
            margin: 0;
            font-size: 1.6rem;
            font-weight: 600;
        }

        .user-info {
            display: flex;
            align-items: center;
            gap: 15px;
        }

        .user-avatar {
            background: var(--primary-blue);
            color: white;
            width: 40px;
            height: 40px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.2rem;
            cursor: pointer;
        }

        /* Content Area */
        .content-area {
            padding: 30px;
        }

        /* Page Header */
        .page-header {
            background: #FF99CC;
            color: white;
            padding: 15px 25px;
            border-radius: 8px 8px 0 0;
            font-size: 1.2rem;
            font-weight: 600;
            margin-bottom: 0;
        }

        /* Review Section */
        .review-section {
            background: white;
            padding: 15px 25px;
            border-bottom: 2px solid #dee2e6;
            font-weight: 600;
            color: #2c3e50;
        }

        /* Submit Button Section */
        .submit-section {
            background: white;
            padding: 20px 25px;
            border-radius: 0 0 8px 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            margin-bottom: 25px;
            display: flex;
            justify-content: flex-end;
        }

        .btn-submit {
            background: var(--green-btn);
            color: white;
            padding: 10px 40px;
            border: none;
            border-radius: 6px;
            font-weight: 600;
            font-size: 1rem;
            transition: all 0.3s;
            cursor: pointer;
        }

        .btn-submit:hover:not(:disabled) {
            background: #218838;
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(40,167,69,0.3);
        }
        
        .btn-submit:disabled {
            background: #6c757d;
            cursor: not-allowed;
        }


        /* Data Table */
        .table-container {
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }

        .table {
            margin: 0;
            font-size: 0.85rem;
        }

        .table thead {
            background: linear-gradient(135deg, var(--primary-blue), #0b5ed7);
            color: white;
        }

        .table thead th {
            padding: 12px 10px;
            font-weight: 600;
            border: none;
            white-space: nowrap;
            vertical-align: middle;
            font-size: 0.85rem;
        }

        .table tbody td {
            padding: 10px;
            vertical-align: middle;
            border-bottom: 1px solid #e9ecef;
            font-size: 0.85rem;
        }

        .table tbody tr:hover {
            background: #f8f9fa;
        }

        .highlight-pink {
            background: var(--pink-highlight) !important;
        }

        /* Checkbox */
        .form-check-input {
            width: 20px;
            height: 20px;
            cursor: pointer;
        }
        /* (เพิ่ม) Loading Overlay */
        .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.6);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 9999;
            backdrop-filter: blur(3px);
        }
        .loading-overlay.active {
            display: flex;
        }
        .loading-content {
            background: white;
            padding: 30px 40px;
            border-radius: 12px;
            text-align: center;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.3);
        }
        .loading-spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #f3f3f3;
            border-top: 4px solid var(--primary-blue);
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 15px;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .loading-text {
            color: #2c3e50;
            font-size: 1.1rem;
            font-weight: 600;
            margin: 0;
        }
        /* Responsive */
        @media (max-width: 768px) {
            .content-area {
                padding: 15px;
            }

            .table {
                font-size: 0.75rem;
            }
        }

        /* Scrollbar */
        .sidebar::-webkit-scrollbar {
            width: 6px;
        }

        .sidebar::-webkit-scrollbar-track {
            background: #1a252f;
        }

        .sidebar::-webkit-scrollbar-thumb {
            background: #34495e;
            border-radius: 3px;
        }

        .table-responsive::-webkit-scrollbar {
            height: 8px;
        }

        .table-responsive::-webkit-scrollbar-track {
            background: #f1f1f1;
        }

        .table-responsive::-webkit-scrollbar-thumb {
            background: #888;
            border-radius: 4px;
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
            <h3><i class="bi bi-building"></i> BMS</h3>
            <button class="close-sidebar" onclick="toggleSidebar()">
                <i class="bi bi-x-lg"></i>
            </button>
        </div>
        <ul class="sidebar-menu">
            <li class="menu-item">
                <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'otbPlan')">
                    <i class="bi bi-clipboard-data"></i>
                    <span>OTB Plan</span>
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
                <h1 class="page-title" id="pageTitle">BMS</h1>
            </div>
            <div class="user-info">
                <span class="d-none d-md-inline">Welcome, Admin</span>
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
                <button type="button" id="btnSyncSAP" class="btn-submit me-3" style="background-color: var(--primary-blue);">
                    <i class="bi bi-arrow-repeat"></i> Sync SAP
                </button>
                <!-- (MODIFIED: Added 'disabled' by default) -->
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
                                <th>Select</th>
                                <th>Year</th>
                                <th>Month</th>
                                <th>Cate</th>
                                <th>Company</th>
                                <th>Segment</th>
                                <th>Brand</th>
                                <th>Vendor</th>
                                <th>Draft PO Date</th>
                                <th>Draft PO/PO no.</th>
                                <th>Draft PO Amount (THB)</th>
                                <th>Draft PO Amount (CCY)</th>
                                <th>Actual PO Amount (THB)</th>
                                <th>Actual Amount (CCY)</th>
                                <th>CCY</th>
                                <th>Acutal Ex. Rate</th>
                                <th>Actual PO Date</th>
                                <th>Actual PO no.</th>
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
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // (MODIFIED: Cache all buttons)
        let btnSyncSAP = document.getElementById("btnSyncSAP");
        let btnSubmit = document.getElementById("btnSubmit");
        let tableBody = document.getElementById("matchTableBody");
        let loadingOverlay = document.getElementById("loadingOverlay");


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

        // (MODIFIED) Function สร้างตาราง
        function buildTable(data) {
            tableBody.innerHTML = ''; // Clear table
            if (!data || data.length === 0) {
                tableBody.innerHTML = '<tr><td colspan="18" class="text-center p-4 text-muted">No data found.</td></tr>';
                btnSubmit.disabled = true; // (MODIFIED: Disable submit if no data)
                return;
            }

            let html = '';
            data.forEach((item, index) => {
                const key = item.Key;
                const draft = item.Draft;
                const actual = item.Actual;

                // (MODIFIED: Set class based on MatchStatus)
                let rowClass = '';
                if (item.MatchStatus === 'Matched') {
                    rowClass = 'highlight-matched'; // Green highlight
                }

                html += `<tr class="${rowClass}">`;

                // (MODIFIED: Add data attributes to checkbox for submission)
                html += `<td>
                            <input type="checkbox" 
                                   class="form-check-input match-checkbox" 
                                   ${item.MatchStatus === 'Matched' ? 'checked' : ''}
                                   data-draft-pos="${draft ? HttpUtility.HtmlAttributeEncode(draft.DraftPONo) : ''}"
                                   data-actual-po="${actual ? HttpUtility.HtmlAttributeEncode(actual.ActualPONo) : ''}"
                            >
                         </td>`;

                // --- Group Keys ---
                html += `<td>${key.Year || ''}</td>`;
                html += `<td>${getMonthName(key.Month) || ''}</td>`;
                html += `<td>${key.Category || ''}</td>`;
                html += `<td>${key.Company || ''}</td>`;
                html += `<td>${key.Segment || ''}</td>`;
                html += `<td>${key.Brand || ''}</td>`;
                html += `<td>${key.Vendor || ''}</td>`;

                // --- Draft PO Data ---
                html += `<td>${draft ? (draft.DraftPODate || '') : ''}</td>`;
                html += `<td style="max-width: 150px; overflow: hidden; text-overflow: ellipsis;" title="${draft ? HttpUtility.HtmlAttributeEncode(draft.DraftPONo) : ''}">${draft ? (draft.DraftPONo || '') : ''}</td>`;
                html += `<td class="text-end">${draft ? formatNumber(draft.DraftAmountTHB) : ''}</td>`;
                html += `<td class="text-end">${draft ? formatNumber(draft.DraftAmountCCY) : ''}</td>`; // (MODIFIED: Re-ordered)

                // --- Actual PO Data ---
                html += `<td class="text-end">${actual ? formatNumber(actual.ActualAmountTHB) : ''}</td>`;
                html += `<td class="text-end">${actual ? formatNumber(actual.ActualAmountCCY) : ''}</td>`; // (MODIFIED: Re-ordered)
                html += `<td>${actual ? (actual.ActualCCY || '') : ''}</td>`;
                html += `<td class="text-end">${actual ? formatNumber(actual.ActualExRate) : ''}</td>`;
                html += `<td>${actual ? (actual.ActualPODate || '') : ''}</td>`;
                html += `<td style="max-width: 150px; overflow: hidden; text-overflow: ellipsis;" title="${actual ? HttpUtility.HtmlAttributeEncode(actual.ActualPONo) : ''}">${actual ? (actual.ActualPONo || '') : ''}</td>`;

                html += '</tr>';
            });
            tableBody.innerHTML = html;

            // (MODIFIED: Enable submit button after loading data)
            btnSubmit.disabled = false;
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
                url: 'Handler/POMatchingHandler.ashx?action=getpo',
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
                    DraftPOs: cb.dataset.draftPos, // "PO-001, PO-002"
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

        let initial = function () {
            // (MODIFIED: Call the refactored sync function)
            btnSyncSAP.addEventListener('click', syncAndLoadData);

            // (MODIFIED: Add click listener for submit)
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