<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="actualPO.aspx.vb" Inherits="BMS.actualPO" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BMS - Draft PO and Actual PO</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <style>
        :root {
            --primary-blue: #0B56A4;
            --sidebar-bg: #2c3e50;
            --sidebar-hover: #34495e;
            --orange-header: #D2691E;
            --light-blue-bg: #E6F2FF;
            --table-header: #4A90E2;
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

        .page-objective {
            background: white;
            padding: 12px 25px;
            border-bottom: 2px solid #dee2e6;
            font-size: 0.9rem;
            color: #495057;
        }

        /* Filter Box */
        .filter-box {
            background: white;
            border-radius: 0 0 8px 8px;
            padding: 25px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            margin-bottom: 25px;
        }

        .filter-title {
            font-size: 1rem;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid var(--primary-blue);
        }

        .form-label {
            font-weight: 600;
            color: #495057;
            margin-bottom: 8px;
            font-size: 0.9rem;
        }

        .form-control, .form-select {
            border: 2px solid #dee2e6;
            padding: 10px 14px;
            border-radius: 6px;
            font-size: 0.95rem;
        }

        .form-control:focus, .form-select:focus {
            border-color: var(--primary-blue);
            box-shadow: 0 0 0 0.2rem rgba(11,86,164,0.15);
        }

        .info-display {
            background: var(--light-blue-bg);
            border: 2px solid #b8d4f1;
            padding: 10px 14px;
            border-radius: 6px;
            font-size: 0.95rem;
            color: #2c3e50;
            min-height: 42px;
            display: flex;
            align-items: center;
            font-weight: 500;
        }

        /* Buttons */
        .btn-custom {
            padding: 10px 25px;
            border-radius: 6px;
            font-weight: 600;
            transition: all 0.3s ease;
            border: none;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            font-size: 0.95rem;
        }

        .btn-clear {
            background: #17a2b8;
            color: white;
        }

        .btn-clear:hover {
            background: #138496;
            transform: translateY(-2px);
        }

        .btn-view {
            background: var(--primary-blue);
            color: white;
        }

        .btn-view:hover {
            background: #094580;
            transform: translateY(-2px);
        }

        .btn-export {
            background: #6c757d;
            color: white;
        }

        .btn-export:hover {
            background: #5a6268;
            transform: translateY(-2px);
        }

        /* Export Button Section */
        .export-section {
            display: flex;
            justify-content: flex-end;
            margin-bottom: 15px;
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

        /* Responsive */
        @media (max-width: 768px) {
            .content-area {
                padding: 15px;
            }

            .table {
                font-size: 0.75rem;
            }

            .filter-box {
                padding: 15px;
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
                    <li><a href="matchActualPO.aspx" class="menu-link">Match Actual PO</a></li>
                    <li><a href="actualPO.aspx" class="menu-link active">Actual PO</a></li>
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
                Actual PO
            </div>
            <!-- Filter Box -->
            <div class="filter-box">
                <div class="filter-title">
                   Search Actual PO
                </div>

                <%-- *** MODIFIED: Added IDs to all controls *** --%>
                <div class="row g-3 mb-3">
                    <div class="col-md-3">
                        <label class="form-label">Year</label>
                        <select id="ddYearFilter" class="form-select">
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Month</label>
                        <select id="ddMonthFilter" class="form-select">
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Company</label>
                        <select id="ddCompanyFilter" class="form-select">
                        </select>
                    </div>
                    <div class="col-md-3">
                    </div>
                </div>

                <div class="row g-3 mb-3">
                    <div class="col-md-6">
                        <label class="form-label">Category</label>
                        <select id="ddCategoryFilter" class="form-select">
                        </select>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label">Segment</label>
                        <select id="ddSegmentFilter" class="form-select">
                        </select>
                    </div>
                </div>

                <div class="row g-3 mb-4">
                    <div class="col-md-6">
                        <label class="form-label">Brand</label>
                        <select id="ddBrandFilter" class="form-select">
                        </select>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label">Vendor</label>
                        <select id="ddVendorFilter" class="form-select">
                        </select>
                    </div>
                </div>

                <!-- Action Buttons -->
                <div class="row">
                    <div class="col-12 text-end">
                        <button type="button" id="btnClearFilter" class="btn btn-clear btn-custom me-2">
                            <i class="bi bi-x-circle"></i> Clear filter
                        </button>
                        <button type="button" id="btnView" class="btn btn-view btn-custom">
                            <i class="bi bi-eye"></i> View
                        </button>
                    </div>
                </div>
            </div>

             <!-- Export Button -->
            <div class="export-section">
                <button type="button" id="btnExport" class="btn btn-export btn-custom">
                    <i class="bi bi-file-earmark-excel"></i> Export TXN
                </button>
            </div>

            <!-- Data Table -->
            <!-- Data Table -->
            <div class="table-container">
                <div class="table-responsive">
                    <table class="table table-hover mb-0">
                        <thead>
                            <tr>
                                <th>Actual PO Date</th>
                                <th>Actual PO no.</th>
                                <th>Type</th>
                                <th>Year</th>
                                <th>Month</th>
                                <th>Category</th>
                                <th>Category name</th>
                                <th>Company</th>
                                <th>Segment</th>
                                <th>Segment name</th>
                                <th>Brand</th>
                                <th>Brand name</th>
                                <th>Vendor</th>
                                <th>Vendor name</th>
                                <th>Amount (THB)</th>
                                <th>Amount (CCY)</th>
                                <th>CCY</th>
                                <th>Ex. Rate</th>
                                <th>Draft PO Ref</th>
                                <th>Remark</th>
                                <th>Status</th>
                                <th>Status date</th>
                            </tr>
                        </thead>
                        <%-- *** MODIFIED: Added ID and placeholder *** --%>
                        <tbody id="actualPOTableBody">
                            <tr>
                                <td colspan="22" class="text-center text-muted p-4">
                                    Please use the filters and click "View" to see data.
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
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script>
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

        // ==========================================
        // ===== NEW: actualPO.aspx SCRIPT LOGIC =====
        // ==========================================

        // --- Cache Filter Elements ---
        let ddYearFilter = document.getElementById('ddYearFilter');
        let ddMonthFilter = document.getElementById('ddMonthFilter');
        let ddCompanyFilter = document.getElementById('ddCompanyFilter');
        let ddCategoryFilter = document.getElementById('ddCategoryFilter');
        let ddSegmentFilter = document.getElementById('ddSegmentFilter');
        let ddBrandFilter = document.getElementById('ddBrandFilter');
        let ddVendorFilter = document.getElementById('ddVendorFilter');
        let btnView = document.getElementById('btnView');
        let btnClearFilter = document.getElementById('btnClearFilter');
        let btnExport = document.getElementById('btnExport');
        let actualPOTableBody = document.getElementById('actualPOTableBody');

        // --- Initializer ---
        function initial() {

            if (ddYearFilter) {
                $(ddYearFilter).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }
            if (ddMonthFilter) {
                $(ddMonthFilter).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }
            if (ddCompanyFilter) {
                $(ddCompanyFilter).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }
            if (ddCategoryFilter) {
                $(ddCategoryFilter).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }
            if (ddSegmentFilter) {
                $(ddSegmentFilter).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }
            if (ddBrandFilter) {
                $(ddBrandFilter).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }
            if (ddVendorFilter) {
                $(ddVendorFilter).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }

            InitMSData();
            
            // Add Event Listeners
            btnView.addEventListener('click', search);
            btnClearFilter.addEventListener('click', clearFilters);
            btnExport.addEventListener('click', exportTXN);
        }

        // --- Clear Filters ---
        function clearFilters() {
            ddYearFilter.value = "";
            ddMonthFilter.value = "";
            ddCompanyFilter.value = "";
            ddCategoryFilter.value = "";
            ddSegmentFilter.value = "";
            ddBrandFilter.value = "";
            ddVendorFilter.value = "";
            InitVendor(ddVendorFilter); // Reset vendor list
            actualPOTableBody.innerHTML = "<tr><td colspan='22' class='text-center text-muted p-4'>Please use the filters and click 'View' to see data.</td></tr>";
        }

        // --- Search Function (AJAX Call) ---
        function search() {
            var formData = new FormData();
            formData.append('year', ddYearFilter.value);
            formData.append('month', ddMonthFilter.value);
            formData.append('company', ddCompanyFilter.value);
            formData.append('category', ddCategoryFilter.value);
            formData.append('segment', ddSegmentFilter.value);
            formData.append('brand', ddBrandFilter.value);
            formData.append('vendor', ddVendorFilter.value);

            // Show loading state
            actualPOTableBody.innerHTML = "<tr><td colspan='22' class='text-center text-muted p-4'><div class='spinner-border spinner-border-sm' role='status'></div> Loading data...</td></tr>";

            $.ajax({
                url: 'Handler/DataPOHandler.ashx?action=getActualPOList', // Call new action
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    actualPOTableBody.innerHTML = response; // Inject HTML
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                    actualPOTableBody.innerHTML = `<tr><td colspan='22' class='text-center text-danger p-4'>Error loading data: ${xhr.responseText}</td></tr>`;
                }
            });
        }

        // --- Export Function ---
        function exportTXN() {
            console.log("Export TXN clicked");
            // Build query string from filters
            var params = new URLSearchParams();
            params.append('action', 'exportActualPO'); // Call new export action
            params.append('year', ddYearFilter.value);
            params.append('month', ddMonthFilter.value);
            params.append('company', ddCompanyFilter.value);
            params.append('category', ddCategoryFilter.value);
            params.append('segment', ddSegmentFilter.value);
            params.append('brand', ddBrandFilter.value);
            params.append('vendor', ddVendorFilter.value);

            // Use window.location to trigger file download (GET request)
            window.location.href = 'Handler/DataPOHandler.ashx?' + params.toString();
        }

        // ==========================================
        // --- Master Data Loaders ---
        // (Copied from other pages for consistency)
        // ==========================================

        function InitMSData() {
            InitSegment(ddSegmentFilter);
            InitCategoty(ddCategoryFilter);
            InitBrand(ddBrandFilter);
            InitVendor(ddVendorFilter);
            InitMSYear(ddYearFilter);
            InitMonth(ddMonthFilter);
            InitCompany(ddCompanyFilter);
            
        }

        function InitSegment(element, addAll = true) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=SegmentMSList',
                type: 'POST',
                data: { addAll: addAll },
                success: (response) => element.innerHTML = response,
                error: (xhr, s, e) => console.log('Error getlist Segment: ' + e)
            });
        }

        function InitMSYear(element, addAll = true) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=YearMSList',
                type: 'POST',
                data: { addAll: addAll },
                success: (response) => element.innerHTML = response,
                error: (xhr, s, e) => console.log('Error getlist Year: ' + e)
            });
        }

        function InitMonth(element, addAll = true) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=MonthMSList',
                type: 'POST',
                data: { addAll: addAll },
                success: (response) => element.innerHTML = response,
                error: (xhr, s, e) => console.log('Error getlist Month: ' + e)
            });
        }
        function InitCompany(element, addAll = true) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=CompanyMSList',
                type: 'POST',
                data: { addAll: addAll },
                success: (response) => element.innerHTML = response,
                error: (xhr, s, e) => console.log('Error getlist Company: ' + e)
            });
        }
        function InitCategoty(element, addAll = true) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=CategoryMSList',
                type: 'POST',
                data: { addAll: addAll },
                success: (response) => element.innerHTML = response,
                error: (xhr, s, e) => console.log('Error getlist Category: ' + e)
            });
        }
        function InitBrand(element, addAll = true) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=BrandMSList',
                type: 'POST',
                data: { addAll: addAll },
                success: (response) => element.innerHTML = response,
                error: (xhr, s, e) => console.log('Error getlist Brand: ' + e)
            });
        }
        function InitVendor(element, addAll = true) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=VendorMSList',
                type: 'POST',
                data: { addAll: addAll },
                success: (response) => element.innerHTML = response,
                error: (xhr, s, e) => console.log('Error getlist Vendor: ' + e)
            });
        }

        // Run initializer on load
        document.addEventListener('DOMContentLoaded', initial);

    </script>
</body>
</html>