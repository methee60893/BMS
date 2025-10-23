<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="transactionOTBSwitching.aspx.vb" Inherits="BMS.transactionOTBSwitching" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BMS - Switching Transaction</title>
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

        /* Two-row header styling */
        .table thead tr:first-child th {
            background: var(--table-header);
            color: white;
            padding: 12px 10px;
            font-weight: 600;
            border: 1px solid #3a7bc8;
            text-align: center;
            vertical-align: middle;
        }

        .table thead tr:nth-child(2) th {
            background: var(--table-header);
            color: white;
            padding: 10px 8px;
            font-weight: 600;
            border: 1px solid #3a7bc8;
            white-space: nowrap;
            font-size: 0.8rem;
        }

        .table tbody td {
            padding: 10px 8px;
            vertical-align: middle;
            border: 1px solid #dee2e6;
            font-size: 0.85rem;
        }

        .table tbody tr:hover {
            background: #f8f9fa;
        }

        .type-switch-in {
            color: #0B56A4;
            font-weight: 600;
        }

        .type-switch-out {
            color: #dc3545;
            font-weight: 600;
        }

        .type-carry-out {
            color: #6c757d;
            font-weight: 600;
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
                    <li><a href="transactionOTBSwitching.aspx" class="menu-link active">Switching Transaction</a></li>
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
                Transaction OTB Switching
            </div>

            <!-- Filter Box -->
            <div class="filter-box">
                <div class="row g-3 mb-3">
                    <div class="col-md-3">
                        <label class="form-label">Type</label>
                        <select class="form-select">
                            <option selected>Switch In-out</option>
                            <option>Switch In</option>
                            <option>Switch Out</option>
                            <option>Carry Out</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Year</label>
                        <select class="form-select">
                            <option selected>2025</option>
                            <option>2024</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Month</label>
                        <select class="form-select">
                            <option>Jan</option>
                            <option>Feb</option>
                            <option>Mar</option>
                            <option>Apr</option>
                            <option>May</option>
                            <option selected>Jun</option>
                            <option>Jul</option>
                            <option>Aug</option>
                            <option>Sep</option>
                            <option>Oct</option>
                            <option>Nov</option>
                            <option>Dec</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Company</label>
                        <select class="form-select">
                            <option selected>KPC</option>
                            <option>KPD</option>
                            <option>KPT</option>
                            <option>KPS</option>
                        </select>
                    </div>
                </div>

                <div class="row g-3 mb-3">
                    <div class="col-md-6">
                        <label class="form-label">Category</label>
                        <select class="form-select">
                            <option selected>221 - FA Leather Goods</option>
                            <option>222 - FA Accessories</option>
                        </select>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label">Segment</label>
                        <select class="form-select">
                            <option selected>O2000 - T/T Normal</option>
                            <option>O3000 - Local Credit</option>
                        </select>
                    </div>
                </div>

                <div class="row g-3 mb-4">
                    <div class="col-md-6">
                        <label class="form-label">Brand</label>
                        <select class="form-select">
                            <option selected>HBS - HUGO BOSS</option>
                            <option>MCM - MCM</option>
                        </select>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label">Vendor</label>
                        <select class="form-select">
                            <option selected>1010900 - HUGO BOSS SOUTH EAST ASIA</option>
                            <option>1011009 - MCM FASHION GROUP LIMITED</option>
                        </select>
                    </div>
                </div>

                <!-- Action Buttons -->
                <div class="row">
                    <div class="col-12 text-end">
                        <button class="btn btn-clear btn-custom me-2">
                            <i class="bi bi-x-circle"></i> Clear Filter
                        </button>
                        <button class="btn btn-view btn-custom">
                            <i class="bi bi-eye"></i> View
                        </button>
                    </div>
                </div>
            </div>

            <!-- Export Button -->
            <div class="export-section">
                <button class="btn btn-export btn-custom">
                    <i class="bi bi-file-earmark-excel"></i> Export TXN
                </button>
            </div>

            <!-- Data Table -->
            <div class="table-container">
                <div class="table-responsive">
                    <table class="table table-bordered mb-0">
                        <thead>
                            <tr>
                                <th rowspan="2">Create date</th>
                                <th rowspan="2">Type</th>
                                <th colspan="11" style="background: #5fa8d3;">From</th>
                                <th rowspan="2">Type</th>
                                <th colspan="11" style="background: #5fa8d3;">To</th>
                                <th rowspan="2">Amount (THB)</th>
                                <th rowspan="2">SAP date</th>
                                <th rowspan="2">Action by</th>
                            </tr>
                            <tr>
                                <!-- From columns -->
                                <th rowspan="1">Year</th>
                                <th rowspan="1">Month</th>
                                <th rowspan="1">Category</th>
                                <th rowspan="1">Category name</th>
                                <th rowspan="1">Company</th>
                                <th rowspan="1">Segment</th>
                                <th rowspan="1">Segment name</th>
                                <th rowspan="1">Brand</th>
                                <th rowspan="1">Brand name</th>
                                <th rowspan="1">Vendor</th>
                                <th rowspan="1">Vendor name</th>
                                <!-- To columns -->
                                <th rowspan="1">Year</th>
                                <th rowspan="1">Month</th>
                                <th rowspan="1">Category</th>
                                <th rowspan="1">Category name</th>
                                <th rowspan="1">Company</th>
                                <th rowspan="1">Segment</th>
                                <th rowspan="1">Segment name</th>
                                <th rowspan="1">Brand</th>
                                <th rowspan="1">Brand name</th>
                                <th rowspan="1">Vendor</th>
                                <th rowspan="1">Vendor name</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>9/6/2025 11:00 AM</td>
                                <td><span class="type-switch-in">Switch in</span></td>
                                <td>2025</td>
                                <td>Jun</td>
                                <td>221</td>
                                <td>FA Leather Goods</td>
                                <td>KPC</td>
                                <td>O3000</td>
                                <td>Local Credit</td>
                                <td>MCM</td>
                                <td>MCM</td>
                                <td>1011009</td>
                                <td>MCM FASHION GROUP LIMITED</td>
                                <td><span class="type-switch-out">Switch out</span></td>
                                <td>2025</td>
                                <td>Jun</td>
                                <td>221</td>
                                <td>FA Leather Goods</td>
                                <td>KPC</td>
                                <td>O3000</td>
                                <td>Local Credit</td>
                                <td>MCM</td>
                                <td>MCM</td>
                                <td>1011009</td>
                                <td>MCM FASHION GROUP LIMITED</td>

                                <td class="text-end">20,000.00</td>
                                <td>9/6/2025 11:00 AM</td>
                                <td>Arunrung T.</td>
                            </tr>
                            <tr>
                                <td>10/6/2025 12:00 AM</td>
                                <td><span class="type-carry-out">Carry out</span></td>
                                <td>2025</td>
                                <td>Jul</td>
                                <td>221</td>
                                <td>FA Leather Goods</td>
                                <td>KPC</td>
                                <td>O3000</td>
                                <td>Local Credit</td>
                                <td>MCM</td>
                                <td>MCM</td>
                                <td>1011009</td>
                                <td>MCM FASHION GROUP LIMITED</td>
                                <td><span class="type-carry-out">Carry out</span></td>
                                <td>2025</td>
                                <td>Jun</td>
                                <td>221</td>
                                <td>FA Leather Goods</td>
                                <td>KPC</td>
                                <td>O3000</td>
                                <td>Local Credit</td>
                                <td>MCM</td>
                                <td>MCM</td>
                                <td>1011009</td>
                                <td>MCM FASHION GROUP LIMITED</td>
                                <td class="text-end">30,000.00</td>
                                <td>10/6/2025 12:00 AM</td>
                                <td>Arunrung T.</td>
                            </tr>
                            <tr>
                                <td>10/6/2025 12:00 AM</td>
                                <td><span class="type-switch-in">Switch in</span></td>
                                <td>2025</td>
                                <td>Jun</td>
                                <td>221</td>
                                <td>FA Leather Goods</td>
                                <td>KPC</td>
                                <td>O3000</td>
                                <td>Local Credit</td>
                                <td>MCM</td>
                                <td>MCM</td>
                                <td>1011009</td>
                                <td>MCM FASHION GROUP LIMITED</td>
                                <td><span class="type-switch-in">Balance in</span></td>
                                <td>2025</td>
                                <td>Jun</td>
                                <td>221</td>
                                <td>FA Leather Goods</td>
                                <td>KPC</td>
                                <td>O3000</td>
                                <td>Local De Ma</td>
                                <td>MCM</td>
                                <td>MCM</td>
                                <td>1011009</td>
                                <td>MCM FASHION GROUP LIMITED</td>
                                <td class="text-end">20,000.00</td>
                                <td>10/6/2025 12:00 AM</td>
                                <td>Arunrung T.</td>
                            </tr>
                            <tr>
                                <td>18/6/2025 12:00 AM</td>
                                <td><span class="type-switch-in">Switch in</span></td>
                                <td>2025</td>
                                <td>Jun</td>
                                <td>221</td>
                                <td>FA Leather Goods</td>
                                <td>KPC</td>
                                <td>O3000</td>
                                <td>Local Credit</td>
                                <td>MCM</td>
                                <td>MCM</td>
                                <td>1011009</td>
                                <td>MCM FASHION GROUP LIMITED</td>
                                <td><span class="type-switch-out">Switch out</span></td>
                                <td>2025</td>
                                <td>Jun</td>
                                <td>221</td>
                                <td>FA Leather Goods</td>
                                <td>KPC</td>
                                <td>O3000</td>
                                <td>Local De Ma</td>
                                <td>MCM</td>
                                <td>MCM</td>
                                <td>1011009</td>
                                <td>MCM FASHION GROUP LIMITED</td>
                                <td class="text-end">15,000.00</td>
                                <td>18/6/2025 11:00 AM</td>
                                <td>Arunrung T.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
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

        // Switch Tab
        function switchTab(tab) {
            const tabs = document.querySelectorAll('.tab-button');
            tabs.forEach(t => t.classList.remove('active'));
            
            const tabContents = document.querySelectorAll('.tab-content');
            tabContents.forEach(tc => tc.classList.remove('active'));
            
            event.target.closest('.tab-button').classList.add('active');
            
            if (tab === 'txn') {
                document.getElementById('txnTab').classList.add('active');
            } else if (tab === 'upload') {
                document.getElementById('uploadTab').classList.add('active');
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
    </script>
