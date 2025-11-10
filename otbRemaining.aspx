<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="otbRemaining.aspx.vb" Inherits="BMS.otbRemaining" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BMS - OTB Remaining</title>
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
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
            background: var(--orange-header);
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
            background: var(--orange-header);
            color: white;
            padding: 15px 25px;
            border-radius: 8px 8px 0 0;
            font-size: 1.2rem;
            font-weight: 600;
            margin-bottom: 0;
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
            background: var(--primary-blue);
            color: white;
        }

        .btn-export:hover {
            background: #094580;
            transform: translateY(-2px);
        }

        /* Detail Summary Box */
        .detail-box {
            background: white;
            border-radius: 8px;
            padding: 25px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            margin-bottom: 25px;
        }

        .detail-title {
            font-size: 1rem;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid var(--primary-blue);
        }

        .detail-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 12px 0;
            border-bottom: 1px solid #e9ecef;
        }

        .detail-row:last-child {
            border-bottom: none;
        }

        .detail-label {
            font-weight: 600;
            color: #495057;
        }

        .detail-value {
            font-weight: 600;
            color: #2c3e50;
            text-align: right;
        }

        .detail-link {
            color: var(--primary-blue);
            text-decoration: none;
            cursor: pointer;
        }

        .detail-link:hover {
            text-decoration: underline;
        }

        /* Data Table */
        .table-container {
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            margin-bottom: 25px;
        }

        .section-title {
            background: #f8f9fa;
            padding: 15px 20px;
            font-weight: 700;
            color: #2c3e50;
            border-bottom: 2px solid #dee2e6;
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

        /* Export Section */
        .export-section {
            display: flex;
            justify-content: center;
            margin: 20px 0;
        }

        /* Responsive */
        @media (max-width: 768px) {
            .content-area {
                padding: 15px;
            }

            .table {
                font-size: 0.75rem;
            }

            .filter-box, .detail-box {
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
                    <li><a href="actualPO.aspx" class="menu-link">Actual PO</a></li>
                </ul>
            </li>
            <li class="menu-item">
                <a href="otbRemaining.aspx" class="menu-link active">
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
                OTB Remaining
            </div>

            <!-- Filter Box -->
            <div class="filter-box">
                <div class="filter-title">
                    OTB Remaining
                </div>

                <div class="row g-3 mb-3">
                    <div class="col-md-3">
                        <label class="form-label">Year</label>
                        <select id="DDYear" class="form-select">
                            </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Month</label>
                        <select id="DDMonth" class="form-select">
                            </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Company</label>
                        <select id="DDCompany" class="form-select">
                            </select>
                    </div>
                    <div class="col-md-3">
                    </div>
                </div>

                <div class="row g-3 mb-3">
                    <div class="col-md-6">
                        <label class="form-label">Category</label>
                        <select id="DDCategory" class="form-select">
                            </select>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label">Segment</label>
                        <select id="DDSegment" class="form-select">
                            </select>
                    </div>
                </div>

                <div class="row g-3 mb-4">
                    <div class="col-md-6">
                        <label class="form-label">Brand</label>
                        <select id="DDBrand" class="form-select">
                            </select>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label">Vendor</label>
                       <select id="DDVendor" class="form-select">
                            </select>
                    </div>
                </div>

                <!-- Action Buttons -->
                <div class="row">
                    <div class="col-12 text-end">
                        <button type="button" id="btnClearFilter" class="btn btn-clear btn-custom me-2">
                            <i class="bi bi-x-circle"></i> Clear filter
                        </button>
                        <button type="button" id="btnView" class="btn btn-view btn-custom me-2">
                            <i class="bi bi-eye"></i> View
                        </button>
                        <button type="button" id="btnTxtSum" class="btn btn-view btn-custom me-2">
                            <i class="bi bi-eye"></i> Export Summary
                        </button>
                    </div>
                </div>
            </div>

            <!-- Detail Summary -->
            <div class="detail-box">
                    <div class="detail-title">Detail</div>
                    
                    <%-- *** MODIFIED: Added IDs to all detail-value divs *** --%>
                    <div class="detail-row">
                        <div class="detail-label">Budget Approved (Original)</div>
                        <div class="detail-value" id="detail_BudgetApproved_Original">
                            0.00 THB
                            <%-- <a href="#" class="detail-link ms-3">Click history</a> --%>
                        </div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Revised Diff</div>
                        <div class="detail-value" id="detail_RevisedDiff">0.00 THB</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Extra</div>
                        <div class="detail-value" id="detail_Budget_Extra">
                            0.00 THB
                            <%-- <a href="#" class="detail-link ms-3">Click history</a> --%>
                        </div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Switch in</div>
                        <div class="detail-value" id="detail_Budget_SwitchIn">
                            0.00 THB
                            <%-- <a href="#" class="detail-link ms-3">Click history</a> --%>
                        </div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Balance in</div>
                        <div class="detail-value" id="detail_Budget_BalanceIn">0.00 THB</div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Carry in</div>
                        <div class="detail-value" id="detail_Budget_CarryIn">0.00 THB</div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Switch to</div>
                        <div class="detail-value" id="detail_Budget_SwitchOut">
                            0.00 THB
                            <%-- <a href="#" class="detail-link ms-3">Click history</a> --%>
                        </div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Balance out</div>
                        <div class="detail-value" id="detail_Budget_BalanceOut">0.00 THB</div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Carry out</div>
                        <div class="detail-value" id="detail_Budget_CarryOut">0.00 THB</div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label"><strong>Total Budget Approved</strong></div>
                        <div class="detail-value"><strong id="detail_TotalBudgetApproved">0.00 THB</strong></div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Actual/Draft</div>
                        <div class="detail-value" id="detail_TotalPO_Usage">
                            0.00 THB
                            <%-- <a href="#" class="detail-link ms-3">Click history</a> --%>
                        </div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label"><strong>Remaining</strong></div>
                        <div class="detail-value"><strong id="detail_Remaining">0.00 THB</strong></div>
                    </div>
                </div>

            <!-- Export Button -->
            <div class="export-section">
                <button class="btn btn-export btn-custom">
                    <i class="bi bi-file-earmark-excel"></i> Export TXN
                </button>
            </div>

            <!-- Other Remaining Table -->
            <div class="table-container">
                <div class="section-title">Other Remaining</div>
                <div class="table-responsive">
                    <table class="table table-hover mb-0">
                        <thead>
                            <tr>
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
                                <th>Total Budget Approved (THB)</th>
                                <th>Draft/Actual PO (THB)</th>
                                <th>Remaining (THB)</th>
                            </tr>
                        </thead>
                        <tbody id="tableViewBody">
                                <tr>
                                    <td colspan="14" class="text-center text-muted p-4">Please select all filters and click "View" to see data.</td>
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
            
            document.querySelectorAll('.menu-link').forEach(link => {
                link.classList.remove('active');
            });
            
            event.currentTarget.classList.add('active');
            document.getElementById('pageTitle').textContent = pageName;
            
            if (window.innerWidth <= 768) {
                toggleSidebar();
            }
        }

        // Close sidebar when clicking outside
        document.addEventListener('click', function(event) {
            const sidebar = document.getElementById('sidebar');
            const menuToggle = document.querySelector('.menu-toggle');
            
            if (!sidebar.contains(event.target) && !menuToggle.contains(event.target)) {
                if (sidebar.classList.contains('active')) {
                    toggleSidebar();
                }
            }
        });

        // ===================================================
        // ===== NEW SCRIPT BLOCK FOR MASTER DATA =========
        // ===================================================

        // Cache elements
        let yearDropdown = document.getElementById("DDYear");
        let monthDropdown = document.getElementById("DDMonth");
        let companyDropdown = document.getElementById("DDCompany");
        let segmentDropdown = document.getElementById("DDSegment");
        let categoryDropdown = document.getElementById("DDCategory");
        let brandDropdown = document.getElementById("DDBrand");
        let vendorDropdown = document.getElementById("DDVendor");
        let btnClearFilter = document.getElementById("btnClearFilter");
        let btnView = document.getElementById("btnView");
        let tableViewBody = document.getElementById("tableViewBody");

        // Initialize
        let initial = function () {
            //InitData master
            InitMSData();
            segmentDropdown.addEventListener('change', changeVendor);
            btnClearFilter.addEventListener('click', function () {
                // Clear filter fields
                yearDropdown.value = "";
                monthDropdown.value = "";
                companyDropdown.value = "";
                segmentDropdown.value = "";
                categoryDropdown.value = "";
                brandDropdown.value = "";
                vendorDropdown.value = "";
                // Re-initialize vendor dropdown (to show all)
                InitVendor(vendorDropdown);
                // Clear table
                tableViewBody.innerHTML = "<tr><td colspan='14' class='text-center text-muted p-4'>Please use the filters and click 'View' to see data.</td></tr>";
            });
            btnView.addEventListener('click', search);
        }

        // Main function to load all master data
        let InitMSData = function () {
            InitSegment(segmentDropdown);
            InitCategoty(categoryDropdown);
            InitBrand(brandDropdown);
            InitVendor(vendorDropdown);
            InitMSYear(yearDropdown);
            InitMonth(monthDropdown);
            InitCompany(companyDropdown);
        }

        // Search function (placeholder - needs handler logic)
        let search = function () {
            var segmentCode = segmentDropdown.value;
            var cate = categoryDropdown.value;
            var brandCode = brandDropdown.value;
            var vendorCode = vendorDropdown.value;
            let OTByear = yearDropdown.value;
            let OTBmonth = monthDropdown.value;
            let OTBcompany = companyDropdown.value;

            var formData = new FormData();
            formData.append('OTByear', OTByear);
            formData.append('OTBmonth', OTBmonth);
            formData.append('OTBCompany', OTBcompany);
            formData.append('OTBCategory', cate);
            formData.append('OTBSegment', segmentCode);
            formData.append('OTBBrand', brandCode);
            formData.append('OTBVendor', vendorCode);

            // Check if all fields are selected
            if (!OTByear) {
                alert("Please select year to view the report.");
                return;
            }

            console.log("Searching with filters...", Object.fromEntries(formData));
            // Show loading state
            tableViewBody.innerHTML = "<tr><td colspan='14' class='text-center text-muted p-4'><div class='spinner-border spinner-border-sm' role='status'></div> Loading data...</td></tr>";
            // Clear detail box
            populateDetailBox(null); // Clear detail box

            $.ajax({
                url: 'Handler/DataOTBRemainingHandler.ashx', // <-- Call the NEW handler
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                dataType: 'json', // Expect JSON back
                success: function (response) {
                    // response will be a DataSet: { "detail": [..], "otherRemaining": [..] }
                    if (response.detail && response.detail.length > 0) {
                        populateDetailBox(response.detail[0]);
                    } else {
                        // No exact match found for detail, but still show "Other"
                        populateDetailBox(null);
                    }

                    if (response.otherRemaining) {
                        populateOtherRemainingTable(response.otherRemaining);
                    } else {
                        // This case shouldn't happen if the SP returns 2 tables, but good to have
                        tableViewBody.innerHTML = "<tr><td colspan='14' class='text-center text-muted p-4'>Error processing response data.</td></tr>";
                    }
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error, xhr.responseText);
                    let errorMsg = "Error loading data.";
                    try {
                        // Try to parse error from our handler
                        let errResponse = JSON.parse(xhr.responseText);
                        if (errResponse.message) errorMsg = errResponse.message;
                    } catch (e) { }

                    tableViewBody.innerHTML = "<tr><td colspan='14' class='text-center text-danger p-4'>" + errorMsg + "</td></tr>";
                    populateDetailBox(null); // Clear detail on error
                }
            });
        }

        // ===== NEW: Helper function to format numbers =====
        function formatTHB(number) {
            if (number === null || number === undefined) {
                return "0.00 THB";
            }
            // Convert to number before formatting
            return parseFloat(number).toLocaleString('en-US', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            }) + " THB";
        }

        // ===== NEW: Helper function to populate detail box =====
        function populateDetailBox(data) {
            if (data) {
                // Populate all fields from the 'detail' (Result Set 1)
                document.getElementById('detail_BudgetApproved_Original').textContent = formatTHB(data.BudgetApproved_Original);
                document.getElementById('detail_RevisedDiff').textContent = formatTHB(data.RevisedDiff); // <-- ADDED
                document.getElementById('detail_Budget_Extra').textContent = formatTHB(data.Budget_Extra);
                document.getElementById('detail_Budget_SwitchIn').textContent = formatTHB(data.Budget_SwitchIn);
                document.getElementById('detail_Budget_BalanceIn').textContent = formatTHB(data.Budget_BalanceIn);
                document.getElementById('detail_Budget_CarryIn').textContent = formatTHB(data.Budget_CarryIn);
                document.getElementById('detail_Budget_SwitchOut').textContent = formatTHB(data.Budget_SwitchOut);
                document.getElementById('detail_Budget_BalanceOut').textContent = formatTHB(data.Budget_BalanceOut);
                document.getElementById('detail_Budget_CarryOut').textContent = formatTHB(data.Budget_CarryOut);
                document.getElementById('detail_TotalBudgetApproved').textContent = formatTHB(data.TotalBudgetApproved);

                // --- Logic from image: TotalPO_Usage = TotalDraftPO + TotalActualPO ---
                // (SP ส่งมา 3 ค่า เราจะใช้ TotalPO_Usage)
                document.getElementById('detail_TotalPO_Usage').textContent = formatTHB(data.TotalPO_Usage);

                document.getElementById('detail_Remaining').textContent = formatTHB(data.Remaining);
            } else {
                // Clear all values to 0.00 THB if no data
                document.getElementById('detail_BudgetApproved_Original').textContent = "0.00 THB";
                document.getElementById('detail_RevisedDiff').textContent = "0.00 THB"; // <-- ADDED
                document.getElementById('detail_Budget_Extra').textContent = "0.00 THB";
                document.getElementById('detail_Budget_SwitchIn').textContent = "0.00 THB";
                document.getElementById('detail_Budget_BalanceIn').textContent = "0.00 THB";
                document.getElementById('detail_Budget_CarryIn').textContent = "0.00 THB";
                document.getElementById('detail_Budget_SwitchOut').textContent = "0.00 THB";
                document.getElementById('detail_Budget_BalanceOut').textContent = "0.00 THB";
                document.getElementById('detail_Budget_CarryOut').textContent = "0.00 THB";
                document.getElementById('detail_TotalBudgetApproved').textContent = "0.00 THB";
                document.getElementById('detail_TotalPO_Usage').textContent = "0.00 THB";
                document.getElementById('detail_Remaining').textContent = "0.00 THB";
            }
        }

        // ===== NEW: Helper function to populate 'Other Remaining' table =====
        function populateOtherRemainingTable(data) {
            if (!data || data.length === 0) {
                tableViewBody.innerHTML = "<tr><td colspan='14' class='text-center text-muted p-4'>No other remaining data found for this period.</td></tr>";
                return;
            }

            let html = "";
            data.forEach(row => {
                // Use .replace(' THB', '') to remove currency symbol for table view
                html += `
                    <tr>
                        <td>${row.Year || ''}</td>
                        <td>${row.MonthName || ''}</td>
                        <td>${row.Category || ''}</td>
                        <td>${row.CategoryName || ''}</td>
                        <td>${row.CompanyName || ''}</td>
                        <td>${row.Segment || ''}</td>
                        <td>${row.SegmentName || ''}</td>
                        <td>${row.Brand || ''}</td>
                        <td>${row.BrandName || ''}</td>
                        <td>${row.Vendor || ''}</td>
                        <td>${row.VendorName || ''}</td>
                        <td class="text-end">${formatTHB(row.TotalBudgetApproved).replace(' THB', '')}</td>
                        <td class="text-end">${formatTHB(row.Draft_Actual_PO_THB).replace(' THB', '')}</td>
                        <td class="text-end">${formatTHB(row.Remaining).replace(' THB', '')}</td>
                    </tr>
                `;
            });
            tableViewBody.innerHTML = html;
        }

        // --- Master Data Loaders (using '...MSList' actions) ---

        let InitSegment = function (segmentDropdown) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=SegmentMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    segmentDropdown.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }

        let InitMSYear = function (yearDropdown) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=YearMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    yearDropdown.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }

        let InitMonth = function (monthDropdown) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=MonthMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    monthDropdown.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }
        let InitCompany = function (companyDropdown) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=CompanyMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    companyDropdown.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }
        let InitCategoty = function (categoryDropdown) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=CategoryMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    categoryDropdown.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }
        let InitBrand = function (brandDropdown) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=BrandMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    brandDropdown.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }
        let InitVendor = function (vendorDropdown) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=VendorMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    vendorDropdown.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }

        // Dependent Dropdown
        let changeVendor = function () {
            var segmentCode = segmentDropdown.value;
            if (!segmentCode) {
                // ถ้าไม่มีค่า ให้โหลด vendor ทั้งหมด
                InitVendor(vendorDropdown);
                return;
            }
            var formData = new FormData();
            formData.append('segmentCode', segmentCode);
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=VendorMSListChg',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    vendorDropdown.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }

        // Initialize on page load
        document.addEventListener('DOMContentLoaded', initial);
    </script>
</body>
</html>