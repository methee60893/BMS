<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="draftOTB.aspx.vb" Inherits="BMS.draftOTB" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BMS - OTB Management System</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
 <style>
        :root {
            --primary-blue: #0d6efd;
            --sidebar-bg: #2c3e50;
            --sidebar-hover: #34495e;
            --orange-header: #ff6b35;
            --green-btn: #28a745;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #ecf0f1;
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
            letter-spacing: 1px;
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
            position: relative;
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

        /* Main Content */
        .main-wrapper {
            width: 100%;
            min-height: 100vh;
            transition: margin-left 0.3s ease;
        }

        /* Top Navigation Bar */
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
            background: #0b5ed7;
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

        .user-info span {
            color: #495057;
            font-weight: 500;
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
            transition: transform 0.3s;
        }

        .user-avatar:hover {
            transform: scale(1.1);
        }

        /* Content Area */
        .content-area {
            padding: 30px;
        }

        /* Filter Box */
        .filter-box {
            background: white;
            border-radius: 10px;
            overflow: hidden;
            margin-bottom: 25px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }

        .filter-header {
            background: var(--orange-header);
            color: white;
            padding: 15px 20px;
            font-weight: 600;
            font-size: 1.1rem;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .filter-body {
            padding: 25px;
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
            transition: all 0.3s;
            font-size: 0.95rem;
        }

        .form-control:focus, .form-select:focus {
            border-color: var(--primary-blue);
            box-shadow: 0 0 0 0.2rem rgba(13,110,253,0.15);
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

        .btn-upload {
            background: var(--green-btn);
            color: white;
        }

        .btn-upload:hover {
            background: #218838;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(40,167,69,0.3);
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
            background: #0b5ed7;
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

        .btn-approve {
            background: var(--green-btn);
            color: white;
        }

        .btn-approve:hover {
            background: #218838;
            transform: translateY(-2px);
        }

        .btn-reject {
            background: #dc3545;
            color: white;
        }

        .btn-reject:hover {
            background: #c82333;
            transform: translateY(-2px);
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
            font-size: 0.9rem;
        }

        .table thead {
            background: linear-gradient(135deg, var(--primary-blue), #0b5ed7);
            color: white;
        }

        .table thead th {
            padding: 14px 12px;
            font-weight: 600;
            border: none;
            white-space: nowrap;
            vertical-align: middle;
        }

        .table tbody td {
            padding: 14px 12px;
            vertical-align: middle;
            border-bottom: 1px solid #e9ecef;
        }

        .table tbody tr {
            transition: background 0.2s;
        }

        .table tbody tr:hover {
            background: #f8f9fa;
        }

        .badge-draft {
            background: #ffc107;
            color: #000;
            padding: 6px 12px;
            border-radius: 4px;
            font-weight: 600;
            font-size: 0.85rem;
        }

        .badge-approved {
            background: #28a745;
            color: white;
            padding: 5px 10px;
            border-radius: 4px;
            font-weight: 600;
            font-size: 0.8rem;
        }

        /* Action Buttons Row */
        .action-buttons {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-top: 20px;
            flex-wrap: wrap;
            gap: 15px;
        }

        .export-buttons {
            display: flex;
            gap: 10px;
        }

        .approval-buttons {
            display: flex;
            gap: 10px;
        }

        /* Responsive */
        @media (max-width: 768px) {
            .content-area {
                padding: 15px;
            }

            .top-navbar {
                padding: 12px 15px;
            }

            .page-title {
                font-size: 1.2rem;
            }

            .table {
                font-size: 0.8rem;
            }

            .filter-body {
                padding: 20px 15px;
            }

            .action-buttons {
                flex-direction: column;
                align-items: stretch;
            }

            .export-buttons, .approval-buttons {
                width: 100%;
                flex-direction: column;
            }

            .btn-custom {
                width: 100%;
                justify-content: center;
            }
        }

        /* Scrollbar Styling */
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

        .sidebar::-webkit-scrollbar-thumb:hover {
            background: #4a5f7f;
        }
    </style>
</head>
<body>
    
    <form id="mainForm" action="/" method="post">

    <!-- Sidebar Overlay -->
    <div class="sidebar-overlay" id="sidebarOverlay" onclick="toggleSidebar()"></div>

    <!-- Sidebar -->
    <div class="sidebar" id="sidebar">
        <div class="sidebar-header">
            <h3><i class="bi bi-building"></i> BMS</h3>
            <button class="close-sidebar" type="button" onclick="toggleSidebar()">
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
                    <li><a href="draftOTB.aspx" class="menu-link active">Draft OTB Plan</a></li>
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
                <button class="menu-toggle" type="button" onclick="toggleSidebar()">
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
            <!-- Upload Card -->
<div class="filter-box mb-3">
    <div class="filter-header">
        <i class="bi bi-cloud-upload"></i>
        Upload File
    </div>
    <div class="filter-body">
        <div class="row align-items-end">
            <div class="col-md-6 col-lg-5">
                <label class="form-label">Select File</label>
                <input type="file" id="fileUpload" class="form-control" accept=".xlsx,.xls,.csv">
            </div>
            <div class="col-md-1 col-lg-1 mt-1 mt-md-0">
                <button id="btnUpload" class="btn btn-upload btn-custom w-100" type="button">
                    <i class="bi bi-upload"></i> Upload
                </button>
            </div>
            <div class="col-md-3 col-lg-5 mt-3 mt-md-0">
                <small class="text-muted">
                    <i class="bi bi-info-circle"></i> Supported formats: Excel (.xlsx, .xls), CSV
                </small>
            </div>
        </div>
    </div>
</div>

<!-- Modal -->
<div class="modal fade" id="previewModal" tabindex="-1" aria-labelledby="previewModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-xl">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="previewModalLabel">Preview Data</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <div id="previewTableContainer"></div>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                <button type="button" id="btnSubmitData" class="btn btn-primary">Submit to Database</button>
            </div>
        </div>
    </div>
</div>

            <!-- Filter Box -->
            <div class="filter-box">
                <div class="filter-header">
                    <i class="bi bi-funnel"></i>
                    Filter Options
                </div>
                <div class="filter-body">
                    <!-- Filter Fields -->
                    <div class="row g-3 mb-3">
                        <div class="col-md-3">
                            <label class="form-label">Type</label>
                            <select id="DDType" class="form-select">
                                <option value="Original" selected>Original</option>
                                <option value="Revise" >Revised</option>
                            </select>
                        </div>
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
                    </div>

                    <div class="row g-3 mb-4">
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
                            <select  id="DDVendor" class="form-select">
                            </select>
                        </div>
                    </div>

                    <!-- Action Buttons -->
                    <div class="row">
                        <div class="col-12 text-end">
                            <button type="button" class="btn btn-clear btn-custom me-2" id="btnClearFilter">
                                <i class="bi bi-x-circle"></i> Clear Filter
                            </button>
                            <button type="button" class="btn btn-view btn-custom" id="btnView">
                                <i class="bi bi-eye"></i> View
                            </button>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Export Buttons -->
            <div class="export-buttons mb-3">
                <button type="button"  class="btn btn-export btn-custom" id="btnExportTXN">
                    <i class="bi bi-file-earmark-excel"></i> Export TXN
                </button>
                <button type="button"  class="btn btn-export btn-custom" id="btnExportSUM">
                    <i class="bi bi-file-earmark-spreadsheet"></i> Export Sum
                </button>
            </div>

            <!-- Data Table -->
            <div class="table-container">
                <div class="table-responsive">
                    <table id="tableView" class="table table-hover mb-0">
                        <thead>
                            <tr>
                                <th style="width: 50px;">
                                
                                </th>
                                <th>Create Date</th>
                                <th>Type</th>
                                <th>Year</th>
                                <th>Month</th>
                                <th>Category</th>
                                <th>Category Name</th>
                                <th>Company</th>
                                <th>Segment</th>
                                <th>Segment Name</th>
                                <th>Brand</th>
                                <th>Brand Name</th>
                                <th>Vendor</th>
                                <th>Vendor Name</th>
                                <th>TO-BE Amount (THB)</th>
                                <th>Current Approved</th>
                                <th>Diff</th>
                                <th>Status</th>
                                <th>Version</th>
                                <th>Remark</th>
                            </tr>
                        </thead>
                        <tbody id="tableViewBody">
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Bottom Action Buttons -->
            <!--<div class="approval-buttons mt-4">
                <button type="button" id="btnApprove" class="btn btn-approve btn-custom">
                    <i class="bi bi-check-circle"></i> Approved
                </button>
                <button type="button" id="btnReject" class="btn btn-reject btn-custom">
                    <i class="bi bi-x-circle"></i> Reject
                </button>
            </div>-->
            <div class="approval-buttons mt-4">
                <button type="button" class="btn btn-success" id="btnApprove">
                    <i class="bi bi-check-circle"></i> Approve Selected
                </button>
                <button type="button" class="btn btn-secondary" id="btnSelectAll">
                    <i class="bi bi-check-all"></i> Select All
                </button>
                <button type="button" class="btn btn-secondary" id="btnDeselectAll">
                    <i class="bi bi-x-circle"></i> Deselect All
                </button>
            </div>
        </div>
    </div>
       </form>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script>
    let mainForm = document.getElementById("mainForm");
    let typeDropdown = document.getElementById("DDType");
    let yearDropdown = document.getElementById("DDYear");
    let monthDropdown = document.getElementById("DDMonth");
    let companyDropdown = document.getElementById("DDCompany");
    let segmentDropdown = document.getElementById("DDSegment");
    let categoryDropdown = document.getElementById("DDCategory");
    let brandDropdown = document.getElementById("DDBrand");
    let vendorDropdown = document.getElementById("DDVendor");
    let btnClearFilter = document.getElementById("btnClearFilter");
    let btnView = document.getElementById("btnView");
    let btnExportTXN = document.getElementById("btnExportTXN");

    let btnApprove = document.getElementById('btnApprove');
    let btnSelectAll = document.getElementById('btnSelectAll');
    let btnDeselectAll = document.getElementById('btnDeselectAll');


    $(document).ready(function () {
        $('#btnUpload').on('click', function (e) {
            e.preventDefault(); // ป้องกัน default behavior (แม้จะเป็น button ก็ตาม)
            console.log("Upload button clicked");

            var fileInput = $('#fileUpload')[0];
            var file = fileInput.files[0];
            var currentUser = '<%= HttpUtility.JavaScriptStringEncode(Session("user").ToString()) %>';
            var uploadBy = currentUser || 'unknown';
            console.log(uploadBy);

            if (!file) {
                alert('Please select a file.');
                return;
            }

            var formData = new FormData();
            formData.append('file', file);
            formData.append('uploadBy', uploadBy); //  ส่งไปกับ request

            $.ajax({
                url: 'Handler/UploadHandler.ashx?action=preview',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    $('#previewTableContainer').html(response);
                    $('#previewModal').modal('show');
                },
                error: function (xhr, status, error) {
                    alert('Error loading preview: ' + error);
                }
            });
        });

        $(document).on('click', '#btnSubmitData', function (e) {
            e.preventDefault(); // ป้องกัน default behavior (แม้จะเป็น button ก็ตาม)
            console.log("Save button clicked");

            if (!confirm('Confirm to save this data to database?')) return;

            var fileInput = $('#fileUpload')[0];
            var file = fileInput.files[0];
            var currentUser = '<%= HttpUtility.JavaScriptStringEncode(Session("user").ToString()) %>';
            var uploadBy = currentUser || 'unknown';
            console.log(uploadBy);

            if (!file) return;

            var formData = new FormData();
            formData.append('file', file);
            formData.append('uploadBy', uploadBy); //  ส่ง uploadBy ไปด้วย

            $.ajax({
                url: 'Handler/UploadHandler.ashx?action=save',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    alert('Data saved successfully!');
                    $('#previewModal').modal('hide');
                    $('#previewTableContainer').empty();
                    $('#fileUpload').val('');
                },
                error: function (xhr, status, error) {
                    alert('Error saving data: ' + error);
                }
            });
        });
    });
        

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
            
            // Toggle submenu
            submenu.classList.toggle('show');
            menuLink.classList.toggle('expanded');
        }

        // Load Page
        function loadPage(event, pageName) {
            event.preventDefault();
            
            // Remove active class from all submenu links
            document.querySelectorAll('.submenu .menu-link').forEach(link => {
                link.classList.remove('active');
            });
            
            // Add active class to clicked link
            event.currentTarget.classList.add('active');
            
            // Update page title
            document.getElementById('pageTitle').textContent = pageName;
            
            // Close sidebar on mobile after selection
            if (window.innerWidth <= 768) {
                toggleSidebar();
            }
            
            console.log('Loading page:', pageName);
            // Here you would implement AJAX call to load page content
            // Example: loadPageContent(pageName);
        }
        // - expand first submenu
        let initial = function () {
            const firstMenuLink = document.querySelector('.menu-link');
            if (firstMenuLink) {
                firstMenuLink.classList.add('expanded');
            }


            //InitData master
            InitMSData();
            segmentDropdown.addEventListener('change', changeVendor);
            btnClearFilter.addEventListener('click', function () {
                mainForm.reset();
                InitMSData();
                tableViewBody.innerHTML = "";
            });
            btnView.addEventListener('click', search);
            btnExportTXN.addEventListener('click', exportTXN);

            // *** ADDED: Approve Button Click Event ***
            btnApprove.addEventListener('click', approveSelectedItems);
    }

    // *** ADDED: Function to Approve Selected Items ***
    let approveSelectedItems = function () {
        let runNosToApprove = [];
        // ค้นหา Checkbox ที่ชื่อ 'checkselect' ที่ถูกเลือก
        $('input[name="checkselect"]:checked').each(function () {
            // ดึงค่า RunNo จาก id (id="checkselect{RunNo}")
            let runNo = this.id.replace('checkselect', '');
            runNosToApprove.push(runNo);
        });

        if (runNosToApprove.length === 0) {
            alert('Please select items to approve.');
            return;
        }

        if (!confirm('Are you sure you want to approve ' + runNosToApprove.length + ' selected items?')) {
            return;
        }

        var currentUser = '<%= HttpUtility.JavaScriptStringEncode(Session("user").ToString()) %>';
            var approvedBy = currentUser || 'unknown';

            var formData = new FormData();
            formData.append('runNos', JSON.stringify(runNosToApprove)); // ส่งเป็น JSON String
            formData.append('approvedBy', approvedBy);

            $.ajax({
                url: 'Handler/DataOTBHandler.ashx?action=approveDraftOTB',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    if (response.trim() === "Success") {
                        alert('Items approved successfully!');
                        search(); // โหลดข้อมูลตารางใหม่
                    } else {
                        alert('Error approving items: ' + response);
                    }
                },
                error: function (xhr, status, error) {
                    console.log('Error approving items: ' + error);
                    alert('An error occurred while approving items.');
                }
            });
        }

        let search = function () {
            var segmentCode = segmentDropdown.value; 
            var cate = categoryDropdown.value;
            var brandCode = brandDropdown.value;
            var vendorCode = vendorDropdown.value;
            var OTBtype = typeDropdown.value;
            let OTByear = yearDropdown.value;
            let OTBmonth = monthDropdown.value;
            let OTBcompany = companyDropdown.value;

            var formData = new FormData();
            formData.append('OTBtype', OTBtype);
            formData.append('OTByear', OTByear);
            formData.append('OTBmonth', OTBmonth);
            formData.append('OTBCompany', OTBcompany);
            formData.append('OTBCategory', cate);
            formData.append('OTBSegment', segmentCode);
            formData.append('OTBBrand', brandCode);
            formData.append('OTBVendor', vendorCode);


            $.ajax({
                url: 'Handler/DataOTBHandler.ashx?action=obtlistbyfilter',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    tableViewBody.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
    }

    let exportTXN = function () {

        console.log("Export TXN clicked");
        // Build query string from filters
        var params = new URLSearchParams();
        params.append('action', 'exportdraftotb');
        params.append('OTBtype', typeDropdown.value);
        params.append('OTByear', yearDropdown.value);
        params.append('OTBmonth', monthDropdown.value);
        params.append('OTBCompany', companyDropdown.value);
        params.append('OTBCategory', categoryDropdown.value);
        params.append('OTBSegment', segmentDropdown.value);
        params.append('OTBBrand', brandDropdown.value);
        params.append('OTBVendor', vendorDropdown.value);

        // Use window.location to trigger file download
        // This is a GET request, so the handler must be adjusted to read from QueryString
        window.location.href = 'Handler/DataOTBHandler.ashx?' + params.toString();
    }

    // Select All Checkbox
    if (btnSelectAll) {
        btnSelectAll.addEventListener('click', function () {
            document.querySelectorAll('input[name="checkselect"]').forEach(cb => {
                cb.checked = true;
            });
        });
    }

    // Deselect All Checkbox
    if (btnDeselectAll) {
        btnDeselectAll.addEventListener('click', function () {
            document.querySelectorAll('input[name="checkselect"]').forEach(cb => {
                cb.checked = false;
            });
        });
    }

    // Approve Button
    if (btnApprove) {
        btnApprove.addEventListener('click', async function () {
            // Get selected checkboxes
            const selectedCheckboxes = document.querySelectorAll('input[name="checkselect"]:checked');

            if (selectedCheckboxes.length === 0) {
                showAlertDraft('warning', 'No Selection', 'Please select at least one record to approve');
                return;
            }

            // Get IDs from checkboxes (assuming ID is in the checkbox id attribute)
            const draftIDs = [];
            selectedCheckboxes.forEach(cb => {
                // Extract ID from checkbox id (e.g., checkselect123 -> 123)
                const id = cb.id.replace('checkselect', '');
                if (id) {
                    draftIDs.push(id);
                }
            });

            if (draftIDs.length === 0) {
                showAlertDraft('warning', 'Error', 'Could not extract draft IDs');
                return;
            }

            // Confirm approval
            if (!confirm(`Are you sure you want to approve ${draftIDs.length} record(s)?`)) {
                return;
            }

            try {
                // Show loading
                showLoadingDraft(true);

                // Prepare form data
                const formData = new FormData();
                formData.append('draftIDs', draftIDs.join(','));
                formData.append('approvedBy', 'Admin'); // TODO: Get from session/user context
                formData.append('remark', ''); // Optional remark

                // Send approval request
                const response = await fetch('Handler/DataOTBHandler.ashx?action=approveDraftOTB', {
                    method: 'POST',
                    body: formData
                });

                const result = await response.json();

                showLoadingDraft(false);

                if (result.success) {
                    showAlertDraft('success', 'Success', result.message);

                    // Refresh the table after 2 seconds
                    setTimeout(() => {
                        search(); // Call existing search function to refresh data
                    }, 2000);
                } else {
                    showAlertDraft('danger', 'Error', result.message);
                }

            } catch (error) {
                showLoadingDraft(false);
                console.error('Approval error:', error);
                showAlertDraft('danger', 'Error', 'Failed to approve records: ' + error.message);
            }
        });
    }

    // Helper function for alerts
    function showAlertDraft(type, title, message) {
        const alertHtml = `
        <div class="alert alert-${type} alert-dismissible fade show" role="alert" style="position: fixed; top: 80px; right: 20px; z-index: 9999; min-width: 300px;">
            <strong>${title}:</strong> ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        </div>
    `;

        // Remove existing alerts
        document.querySelectorAll('.alert').forEach(el => {
            if (el.style.position === 'fixed') el.remove();
        });

        // Add new alert
        document.body.insertAdjacentHTML('beforeend', alertHtml);

        // Auto dismiss after 5 seconds
        setTimeout(() => {
            document.querySelectorAll('.alert[style*="position: fixed"]').forEach(el => {
                el.classList.remove('show');
                setTimeout(() => el.remove(), 150);
            });
        }, 5000);
    }

    // Helper function for loading overlay
    function showLoadingDraft(show) {
        const loadingHtml = `
        <div id="loadingOverlayDraft" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 9999; display: flex; align-items: center; justify-content: center;">
            <div class="spinner-border text-light" role="status" style="width: 3rem; height: 3rem;">
                <span class="visually-hidden">Loading...</span>
            </div>
        </div>
    `;

        if (show) {
            document.body.insertAdjacentHTML('beforeend', loadingHtml);
        } else {
            const overlay = document.getElementById('loadingOverlayDraft');
            if (overlay) overlay.remove();
        }
    }

        let InitMSData = function () {
            InitSegment(segmentDropdown);
            InitCategoty(categoryDropdown);
            InitBrand(brandDropdown);
            InitVendor(vendorDropdown);
            InitMSYear(yearDropdown);
            InitMonth(monthDropdown);
            InitCompany(companyDropdown);
        }

        let InitSegment = function (segmentDropdown) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=SegmentList',
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
                url: 'Handler/MasterDataHandler.ashx?action=YearList',
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
                url: 'Handler/MasterDataHandler.ashx?action=MonthList',
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
            // Implement month initialization if needed
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=CompanyList',
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
                url: 'Handler/MasterDataHandler.ashx?action=CategoryList',
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
                url: 'Handler/MasterDataHandler.ashx?action=BrandList',
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
                url: 'Handler/MasterDataHandler.ashx?action=VendorList',
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
                url: 'Handler/MasterDataHandler.ashx?action=VendorListChg',
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

        // Close sidebar when clicking outside on mobile
        document.addEventListener('click', function(event) {
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
            