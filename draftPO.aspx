<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="draftPO.aspx.vb" Inherits="BMS.draftPO" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BMS - Draft PO</title>
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
            --yellow-btn: #FFC107;
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

        .btn-action {
            background: var(--yellow-btn);
            color: #000;
            padding: 6px 15px;
            font-size: 0.85rem;
        }

        .btn-action:hover {
            background: #e0a800;
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

        .status-cancelled {
            color: #dc3545;
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
                    <li><a href="draftPO.aspx" class="menu-link active">Draft PO</a></li>
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
                Draft PO
            </div>

            <!-- Filter Box -->
            <div class="filter-box">
                <div class="filter-title">
                    Search Draft PO
                </div>

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
                            <i class="bi bi-x-circle"></i> Clear Filter
                        </button>
                        <button type="button" id="btnView" class="btn btn-view btn-custom">
                            <i class="bi bi-eye"></i> View
                        </button>
                    </div>
                </div>
            </div>

            <!-- Export Button -->
            <div class="export-section">
                <button id="btnExport" class="btn btn-export btn-custom">
                    <i class="bi bi-file-earmark-excel"></i> Export TXN
                </button>
            </div>

            <!-- Data Table -->
            <div class="table-container">
                <div class="table-responsive">
                    <table class="table table-hover mb-0">
                        <thead>
                            <tr>
                                <th>Draft PO Date</th>
                                <th>Draft PO no.</th>
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
                                <th>Actual PO Ref</th>
                                <th>Status</th>
                                <th>Status date</th>
                                <th>Remark</th>
                                <th>Action by</th>
                                <th>Action</th>
                            </tr>
                        </thead>
                        <tbody id="draftPOTableBody">
                            <tr>
                                <td colspan="24" class="text-center text-muted p-4">
                                    Please use the filters and click "View" to see data.
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <!-- ================================================= -->
    <!-- ===== ADDED: EDIT PO MODAL ===== -->
    <!-- ================================================= -->
    <div class="modal fade" id="editPOModal" tabindex="-1" aria-labelledby="editPOModalLabel" data-bs-backdrop="static" data-bs-keyboard="false">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header" style="background-color: var(--yellow-btn); color: #333;">
                    <h5 class="modal-title" id="editPOModalLabel">
                        <i class="bi bi-pencil-fill"></i> Edit Draft PO
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div id="editPOContainer">
                        <div class="form-container" style="padding: 15px; box-shadow: none;">
                            <div class="form-section" style="border: none; padding: 0;">
                                <!-- Hidden field for ID -->
                                <input type="hidden" id="hdnDraftPOID">

                                <!-- Row 1 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label for="ddYearEdit">Year</label>
                                        <select id="ddYearEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="year"></div>
                                    </div>
                                    <div class="form-group">
                                        <label for="ddCategoryEdit">Category</label>
                                        <select id="ddCategoryEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="category"></div>
                                    </div>
                                </div>

                                <!-- Row 2 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label for="ddMonthEdit">Month</label>
                                        <select id="ddMonthEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="month"></div>
                                    </div>
                                    <div class="form-group">
                                        <label for="ddSegmentEdit">Segment</label>
                                        <select id="ddSegmentEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="segment"></div>
                                    </div>
                                </div>

                                <!-- Row 3 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label for="ddCompanyEdit">Company</label>
                                        <select id="ddCompanyEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="company"></div>
                                    </div>
                                    <div class="form-group">
                                        <label for="ddBrandEdit">Brand</label>
                                        <select id="ddBrandEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="brand"></div>
                                    </div>
                                </div>

                                <!-- Row 4 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label for="txtPONOEdit">Draft PO no. (Readonly)</label>
                                        <input id="txtPONOEdit" type="text" class="form-control" readonly style="background: #e9ecef;" autocomplete="off">
                                        <div class="validation-message" data-field="pono"></div>
                                    </div>
                                    <div class="form-group">
                                        <label for="ddVendorEdit">Vendor</label>
                                        <select id="ddVendorEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="vendor"></div>
                                    </div>
                                </div>

                                <!-- Row 5 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label for="txtAmtCCYEdit">Amount (CCY)</label>
                                        <input id="txtAmtCCYEdit" type="text" class="form-control" placeholder="0.00" autocomplete="off">
                                        <div class="validation-message" data-field="amtCCY"></div>
                                    </div>
                                    <div class="form-group">
                                        <label for="ddCCYEdit">CCY</label>
                                        <select id="ddCCYEdit" class="form-select">
                                            <option value="">-- Select CCY --</option>
                                            <option>USD</option>
                                            <option>THB</option>
                                            <option>EUR</option>
                                            <option>JPY</option>
                                            <option>SGD</option>
                                        </select>
                                        <div class="validation-message" data-field="ccy"></div>
                                    </div>
                                </div>

                                <!-- Row 6 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label for="txtExRateEdit">Exchange rate</label>
                                        <input id="txtExRateEdit" type="text" class="form-control" placeholder="0.00" autocomplete="off">
                                        <div class="validation-message" data-field="exRate"></div>
                                    </div>
                                    <div class="form-group">
                                        <label for="txtAmtTHBEdit">Amount (THB)</label>
                                        <input id="txtAmtTHBEdit" type="text" class="form-control" readonly style="background: #e9ecef;" autocomplete="off">
                                    </div>
                                </div>
                                
                                <!-- Row 7 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group" style="flex: 1;">
                                        <label for="txtRemarkEdit">Remark</label>
                                        <input id="txtRemarkEdit" type="text" class="form-control" placeholder="Enter remark" autocomplete="off">
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">
                        <i class="bi bi-x-circle"></i> Cancel
                    </button>
                    <button type="button" id="btnSaveChanges" class="btn btn-success" >
                        <i class="bi bi-check-circle-fill"></i> Save Changes
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- ================================================= -->
    <!-- ===== ADDED: HELPER MODALS (Error, Success, Loading) ===== -->
    <!-- ================================================= -->
    <div class="modal fade" id="errorValidationModal" tabindex="-1" data-bs-backdrop="static" data-bs-keyboard="false">
        <div class="modal-dialog modal-dialog-centered modal-lg">
            <div class="modal-content">
                <div class="modal-header bg-danger text-white">
                    <h5 class="modal-title"><i class="bi bi-exclamation-triangle-fill me-2"></i>Validation Error</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="alert alert-danger d-flex align-items-center mb-3" role="alert">
                        <i class="bi bi-x-circle-fill fs-4 me-3"></i>
                        <div>
                            <strong>Please correct the following errors:</strong>
                            <p class="mb-0 mt-1" id="errorSummaryText">Some fields require your attention</p>
                        </div>
                    </div>
                    <div id="errorListContainer" class="error-list-container"></div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">
                        <i class="bi bi-arrow-left me-2"></i>Go Back to Fix
                    </button>
                </div>
            </div>
        </div>
    </div>
    
    <div class="modal fade" id="successModal" tabindex="-1" data-bs-backdrop="static" data-bs-keyboard="false">
        <div class="modal-dialog modal-dialog-centered">
            <div class="modal-content">
                <div class="modal-body text-center p-4">
                    <i class="bi bi-check-circle-fill text-success" style="font-size: 4rem; margin-bottom: 1rem;"></i>
                    <h4 class="modal-title mb-2" id="successModalTitle">Success!</h4>
                    <p id="successModalMessage">Your changes have been saved.</p>
                    <button type="button" class="btn btn-success mt-2" data-bs-dismiss="modal" style="min-width: 100px;">
                        OK
                    </button>
                </div>
            </div>
        </div>
    </div>
    <!-- (Loading overlay will be created by JS) -->
    
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script>

        // ==========================================
        // Global Variables
        // ==========================================
        let editPOModal, errorValidationModal, successModal;

        // Filter Elements
        let ddYearFilter, ddMonthFilter, ddCompanyFilter, ddCategoryFilter, ddSegmentFilter, ddBrandFilter, ddVendorFilter;
        let btnView, btnClearFilter, btnExport;
        let draftPOTableBody;

        // Edit Modal Elements
        let hdnDraftPOID, ddYearEdit, ddMonthEdit, ddCompanyEdit, ddCategoryEdit, ddSegmentEdit, ddBrandEdit, ddVendorEdit;
        let txtPONOEdit, txtAmtCCYEdit, ddCCYEdit, txtExRateEdit, txtAmtTHBEdit, txtRemarkEdit;
        let btnSaveChanges;

        // ==========================================
        // Initializer
        // ==========================================
        function initial() {
            // Cache Filter Elements
            ddYearFilter = document.getElementById('ddYearFilter');
            ddMonthFilter = document.getElementById('ddMonthFilter');
            ddCompanyFilter = document.getElementById('ddCompanyFilter');
            ddCategoryFilter = document.getElementById('ddCategoryFilter');
            ddSegmentFilter = document.getElementById('ddSegmentFilter');
            ddBrandFilter = document.getElementById('ddBrandFilter');
            ddVendorFilter = document.getElementById('ddVendorFilter');
            btnView = document.getElementById('btnView');
            btnClearFilter = document.getElementById('btnClearFilter');
            btnExport = document.getElementById('btnExport');
            draftPOTableBody = document.getElementById('draftPOTableBody');

            // Cache Edit Modal Elements
            hdnDraftPOID = document.getElementById('hdnDraftPOID');
            ddYearEdit = document.getElementById('ddYearEdit');
            ddMonthEdit = document.getElementById('ddMonthEdit');
            ddCompanyEdit = document.getElementById('ddCompanyEdit');
            ddCategoryEdit = document.getElementById('ddCategoryEdit');
            ddSegmentEdit = document.getElementById('ddSegmentEdit');
            ddBrandEdit = document.getElementById('ddBrandEdit');
            ddVendorEdit = document.getElementById('ddVendorEdit');
            txtPONOEdit = document.getElementById('txtPONOEdit');
            txtAmtCCYEdit = document.getElementById('txtAmtCCYEdit');
            ddCCYEdit = document.getElementById('ddCCYEdit');
            txtExRateEdit = document.getElementById('txtExRateEdit');
            txtAmtTHBEdit = document.getElementById('txtAmtTHBEdit');
            txtRemarkEdit = document.getElementById('txtRemarkEdit');
            btnSaveChanges = document.getElementById('btnSaveChanges');

            // Init Modals
            editPOModal = new bootstrap.Modal(document.getElementById('editPOModal'));
            errorValidationModal = new bootstrap.Modal(document.getElementById('errorValidationModal'));
            successModal = new bootstrap.Modal(document.getElementById('successModal'));

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

            if (ddCCYEdit) {
                $(ddCCYEdit).select2({
                    theme: "bootstrap-5",
                    dropdownParent: $("#editPOModal"), 
                    allowClear: true
                });
            }
            if (ddYearEdit) {
                $(ddYearEdit).select2({
                    theme: "bootstrap-5",
                    dropdownParent: $("#editPOModal"), 
                    allowClear: true
                });
            }
            if (ddMonthEdit) {
                $(ddMonthEdit).select2({
                    theme: "bootstrap-5",
                    dropdownParent: $("#editPOModal"), 
                    allowClear: true
                });
            }
            if (ddCompanyEdit) {
                $(ddCompanyEdit).select2({
                    theme: "bootstrap-5",
                    dropdownParent: $("#editPOModal"), 
                    allowClear: true
                });
            }
            if (ddCategoryEdit) {
                $(ddCategoryEdit).select2({
                    theme: "bootstrap-5",
                    dropdownParent: $("#editPOModal"), 
                    allowClear: true
                });
            }
            if (ddSegmentEdit) {
                $(ddSegmentEdit).select2({
                    theme: "bootstrap-5",
                    dropdownParent: $("#editPOModal"), 
                    allowClear: true
                });
            }
            if (ddBrandEdit) {
                $(ddBrandEdit).select2({
                    theme: "bootstrap-5",
                    dropdownParent: $("#editPOModal"), 
                    allowClear: true
                });
            }
            if (ddVendorEdit) {
                $(ddVendorEdit).select2({
                    theme: "bootstrap-5",
                    dropdownParent: $("#editPOModal"), 
                    allowClear: true
                });
            }

            // Load Master Data into Filters
            InitMSDataFilters();
            // Load Master Data into Edit Modal (run once)
            InitMSDataEditModal();

            // Add Event Listeners
            btnView.addEventListener('click', handleViewData);
            btnClearFilter.addEventListener('click', clearFilters);
            btnSaveChanges.addEventListener('click', handleSaveChanges);

            // Add Listeners for Edit Modal Calculations
            txtAmtCCYEdit.addEventListener('input', currencyCalEdit);
            txtExRateEdit.addEventListener('input', currencyCalEdit);
            ddCCYEdit.addEventListener('change', currencyCalEdit);
        }

        // ==========================================
        // Master Data Loading
        // ==========================================

        function InitMSDataFilters() {
            // Load all master data into the filter dropdowns
            InitMSYear(ddYearFilter);
            InitMonth(ddMonthFilter);
            InitCompany(ddCompanyFilter);
            InitCategoty(ddCategoryFilter);
            InitSegment(ddSegmentFilter);
            InitBrand(ddBrandFilter);
            InitVendor(ddVendorFilter);
        }

        function InitMSDataEditModal() {
            // Pre-load master data into the edit modal dropdowns
            InitMSYear(ddYearEdit, false); // false = don't add "All"
            InitMonth(ddMonthEdit, false);
            InitCompany(ddCompanyEdit, false);
            InitCategoty(ddCategoryEdit, false);
            InitSegment(ddSegmentEdit, false);
            InitBrand(ddBrandEdit, false);
            InitVendor(ddVendorEdit, false);
        }

        // Generic Master Data Functions
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

        // ==========================================
        // Filter & View Data Functions
        // ==========================================

        function clearFilters() {
            ddYearFilter.value = "";
            ddMonthFilter.value = "";
            ddCompanyFilter.value = "";
            ddCategoryFilter.value = "";
            ddSegmentFilter.value = "";
            ddBrandFilter.value = "";
            ddVendorFilter.value = "";
        }

        async function handleViewData() {
            showLoading(true, "Fetching data...");

            const formData = new FormData();
            formData.append('year', ddYearFilter.value);
            formData.append('month', ddMonthFilter.value);
            formData.append('company', ddCompanyFilter.value);
            formData.append('category', ddCategoryFilter.value);
            formData.append('segment', ddSegmentFilter.value);
            formData.append('brand', ddBrandFilter.value);
            formData.append('vendor', ddVendorFilter.value);

            try {
                const response = await fetch('Handler/DataPOHandler.ashx?action=getDraftPOList', {
                    method: 'POST',
                    body: formData
                });
                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

                const data = await response.json();
                populateTable(data);

            } catch (error) {
                console.error('Fetch error:', error);
                draftPOTableBody.innerHTML = `<tr><td colspan="24" class="text-center text-danger p-4">Error loading data: ${error.message}</td></tr>`;
            } finally {
                showLoading(false);
            }
        }

        function populateTable(data) {
            draftPOTableBody.innerHTML = ""; // Clear existing data

            if (!data || data.length === 0) {
                draftPOTableBody.innerHTML = `<tr><td colspan="24" class="text-center text-muted p-4">No Draft PO records found for the selected filters.</td></tr>`;
                return;
            }

            const rowsHtml = data.map(row => {
                const createdDate = row.Created_Date ? new Date(row.Created_Date).toLocaleString('en-GB') : '';
                const statusDate = row.Status_Date ? new Date(row.Status_Date).toLocaleString('en-GB') : '';
                const statusClass = row.Status?.toLowerCase() === 'cancelled' ? 'status-cancelled' : '';

                // Helper to format numbers
                const formatNum = (num, decimals = 2) => (num != null ? parseFloat(num).toFixed(decimals) : '0.00');

                return `
                    <tr>
                        <td>${createdDate}</td>
                        <td>${row.DraftPO_No || ''}</td>
                        <td>${row.PO_Type || ''}</td>
                        <td>${row.PO_Year || ''}</td>
                        <td>${row.PO_Month_Name || ''}</td>
                        <td>${row.Category_Code || ''}</td>
                        <td>${row.Category_Name || ''}</td>
                        <td>${row.Company_Code || ''}</td>
                        <td>${row.Segment_Code || ''}</td>
                        <td>${row.Segment_Name || ''}</td>
                        <td>${row.Brand_Code || ''}</td>
                        <td>${row.Brand_Name || ''}</td>
                        <td>${row.Vendor_Code || ''}</td>
                        <td>${row.Vendor_Name || ''}</td>
                        <td class="text-end">${formatNum(row.Amount_THB)}</td>
                        <td class="text-end">${formatNum(row.Amount_CCY)}</td>
                        <td>${row.CCY || ''}</td>
                        <td class="text-end">${formatNum(row.Exchange_Rate, 4)}</td>
                        <td>${row.Actual_PO_Ref || ''}</td>
                        <td class="${statusClass}">${row.Status || ''}</td>
                        <td>${statusDate}</td>
                        <td>${row.Remark || ''}</td>
                        <td>${row.Status_By || ''}</td>
                        <td>
                            <button class="btn btn-action btn-edit-po" 
                                    data-draftpoid="${row.DraftPO_ID}" 
                                    title="Edit ${row.DraftPO_No}">
                                <i class="bi bi-pencil"></i> Edit
                            </button>
                        </td>
                    </tr>
                `;
            }).join('');

            draftPOTableBody.innerHTML = rowsHtml;
            addEditButtonListeners();
        }

        // ==========================================
        // Edit PO Functions
        // ==========================================

        function addEditButtonListeners() {
            document.querySelectorAll('.btn-edit-po').forEach(button => {
                button.addEventListener('click', handleEditClick);
            });
        }

        async function handleEditClick(event) {
            const draftPOID = event.currentTarget.dataset.draftpoid;
            if (!draftPOID) return;

            showLoading(true, "Loading details...");
            clearValidationErrors('editPOModal'); // Clear errors in modal

            const formData = new FormData();
            formData.append('draftPOID', draftPOID);

            try {
                const response = await fetch('Handler/DataPOHandler.ashx?action=getDraftPODetails', {
                    method: 'POST',
                    body: formData
                });
                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

                const data = await response.json();
                populateEditModal(data);
                editPOModal.show();

            } catch (error) {
                console.error('Fetch details error:', error);
                showErrorModal({ 'general': 'Error loading details: ' + error.message });
            } finally {
                showLoading(false);
            }
        }

        function populateEditModal(data) {
            hdnDraftPOID.value = data.DraftPO_ID;
            ddYearEdit.value = data.PO_Year;
            ddMonthEdit.value = data.PO_Month;
            ddCompanyEdit.value = data.Company_Code;
            ddCategoryEdit.value = data.Category_Code;
            ddSegmentEdit.value = data.Segment_Code;
            ddBrandEdit.value = data.Brand_Code;
            ddVendorEdit.value = data.Vendor_Code;
            txtPONOEdit.value = data.DraftPO_No;
            txtAmtCCYEdit.value = parseFloat(data.Amount_CCY || 0).toFixed(2);
            ddCCYEdit.value = data.CCY;
            txtExRateEdit.value = parseFloat(data.Exchange_Rate || 0).toFixed(4);
            txtAmtTHBEdit.value = parseFloat(data.Amount_THB || 0).toFixed(2);
            txtRemarkEdit.value = data.Remark || '';

            // Trigger calculation in case CCY is THB
            currencyCalEdit();
        }

        // Calculation for Edit Modal
        function currencyCalEdit() {
            let amtCCY = parseFloat(txtAmtCCYEdit.value) || 0;
            let exRate = parseFloat(txtExRateEdit.value) || 0;

            if (ddCCYEdit.value === 'THB') {
                txtExRateEdit.value = "1.0000";
                exRate = 1.00;
                txtExRateEdit.readOnly = true;
            } else {
                txtExRateEdit.readOnly = false;
            }
            txtAmtTHBEdit.value = (amtCCY * exRate).toFixed(2);
        }

        // Save Changes
        async function handleSaveChanges() {
            clearValidationErrors('editPOModal');
            showLoading(true, "Validating...");

            // 1. Validate
            const formData = new FormData();
            formData.append('draftPOID', hdnDraftPOID.value);
            formData.append('year', ddYearEdit.value);
            formData.append('month', ddMonthEdit.value);
            formData.append('company', ddCompanyEdit.value);
            formData.append('category', ddCategoryEdit.value);
            formData.append('segment', ddSegmentEdit.value);
            formData.append('brand', ddBrandEdit.value);
            formData.append('vendor', ddVendorEdit.value);
            formData.append('pono', txtPONOEdit.value); // (Readonly)
            formData.append('amtCCY', txtAmtCCYEdit.value);
            formData.append('ccy', ddCCYEdit.value);
            formData.append('exRate', txtExRateEdit.value);
            formData.append('remark', txtRemarkEdit.value);

            try {
                const validationResponse = await fetch('Handler/ValidateHandler.ashx?action=validateDraftPOEdit', {
                    method: 'POST',
                    body: formData
                });
                if (!validationResponse.ok) throw new Error(`Validation HTTP error! status: ${validationResponse.status}`);

                const validationResult = await validationResponse.json();

                if (!validationResult.success) {
                    showLoading(false);
                    showErrorModal(validationResult.errors, 'Edit Draft PO', 'editPOModal');
                    showValidationErrorsOnForm(validationResult.errors, 'editPOModal');
                    return; // Stop execution
                }

                // 2. Save
                showLoading(true, "Saving changes...");

                const saveResponse = await fetch('Handler/DataPOHandler.ashx?action=saveDraftPOEdit', {
                    method: 'POST',
                    body: formData
                });
                if (!saveResponse.ok) throw new Error(`Save HTTP error! status: ${saveResponse.status}`);

                const saveResult = await saveResponse.json();
                showLoading(false);

                if (saveResult.success) {
                    editPOModal.hide();
                    successModal.show();
                    handleViewData(); // Refresh the table
                } else {
                    // Show save error
                    showErrorModal({ 'general': 'Save Failed: ' + saveResult.message }, 'Edit Draft PO');
                }

            } catch (error) {
                showLoading(false);
                console.error('Save changes error:', error);
                showErrorModal({ 'general': 'A system error occurred: ' + error.message });
            }
        }


        // ==========================================
        // Helper Functions (Modals, Validation)
        // ==========================================

        function showLoading(show, text = "Loading...") {
            let overlay = document.getElementById('loadingOverlay');
            if (show) {
                if (!overlay) {
                    const loadingHtml = `
                        <div id="loadingOverlay" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 10000; display: flex; flex-direction: column; align-items: center; justify-content: center; transition: opacity 0.3s;">
                            <div class="spinner-border text-light" role="status" style="width: 3rem; height: 3rem;">
                                <span class="visually-hidden">Loading...</span>
                            </div>
                            <p class="text-light mt-3 mb-0" id="loadingOverlayText" style="font-size: 1.1rem;">${text}</p>
                        </div>`;
                    document.body.insertAdjacentHTML('beforeend', loadingHtml);
                }
                document.getElementById('loadingOverlayText').textContent = text;
            } else {
                if (overlay) {
                    overlay.style.opacity = '0';
                    setTimeout(() => overlay.remove(), 300);
                }
            }
        }

        function clearValidationErrors(modalId = null) {
            const scope = modalId ? document.getElementById(modalId) : document;
            if (!scope) return;

            scope.querySelectorAll('.has-error').forEach(el => el.classList.remove('has-error'));
            scope.querySelectorAll('.validation-message').forEach(el => {
                el.textContent = '';
                el.style.display = 'none';
            });
        }

        function showValidationErrorsOnForm(errors, modalId = null) {
            const scope = modalId ? document.getElementById(modalId) : document;
            if (!scope) return;

            const idPrefixes = modalId ? { txt: 'txt', dd: 'dd', suffix: 'Edit' } : { txt: 'txt', dd: 'DD', suffix: '' };

            for (const field in errors) {
                const elId = field === 'pono' ? `${idPrefixes.txt}PONO${idPrefixes.suffix}` :
                    field === 'amtCCY' ? `${idPrefixes.txt}AmtCCY${idPrefixes.suffix}` :
                        field === 'exRate' ? `${idPrefixes.txt}ExRate${idPrefixes.suffix}` :
                            `${idPrefixes.dd}${field.charAt(0).toUpperCase() + field.slice(1)}${idPrefixes.suffix}`;

                const el = scope.querySelector(`#${elId}`);
                const msgEl = scope.querySelector(`.validation-message[data-field="${field}"]`);

                if (el) el.classList.add('has-error');
                if (msgEl) {
                    msgEl.textContent = errors[field];
                    msgEl.style.display = 'block';
                }
            }
        }

        function showErrorModal(errors, transactionType = 'Draft PO', modalId = null) {
            const errorCount = Object.keys(errors).length;
            const summaryText = document.getElementById('errorSummaryText');
            summaryText.innerHTML = `Found ${errorCount} validation error${errorCount > 1 ? 's' : ''} in your ${transactionType}.`;

            const errorListContainer = document.getElementById('errorListContainer');
            let errorHtml = '';

            const fieldInfo = {
                'year': 'Year', 'month': 'Month', 'company': 'Company', 'category': 'Category',
                'segment': 'Segment', 'brand': 'Brand', 'vendor': 'Vendor', 'pono': 'Draft PO No.',
                'amtCCY': 'Amount (CCY)', 'ccy': 'CCY', 'exRate': 'Exchange rate',
                'general': 'General Error'
            };

            const idPrefixes = modalId ? { txt: 'txt', dd: 'dd', suffix: 'Edit' } : { txt: 'txt', dd: 'DD', suffix: '' };

            for (const field in errors) {
                const message = errors[field];
                const fieldName = fieldInfo[field] || field;
                const fieldId = field === 'pono' ? `${idPrefixes.txt}PONO${idPrefixes.suffix}` :
                    field === 'amtCCY' ? `${idPrefixes.txt}AmtCCY${idPrefixes.suffix}` :
                        field === 'exRate' ? `${idPrefixes.txt}ExRate${idPrefixes.suffix}` :
                            `${idPrefixes.dd}${field.charAt(0).toUpperCase() + field.slice(1)}${idPrefixes.suffix}`;

                errorHtml += `
                    <div class="error-item" ${field !== 'general' ? `data-field="${fieldId}" data-modal="${modalId || ''}"` : ''}>
                        <div class="error-item-icon"><i class="bi bi-x-circle"></i></div>
                        <div class="error-item-content">
                            <div class="error-field-name">
                                <span class="error-section-badge">${fieldName}</span>
                            </div>
                            <p class="error-message">${message}</p>
                        </div>
                    </div>`;
            }

            errorListContainer.innerHTML = errorHtml;
            errorValidationModal.show();

            // Add click-to-focus behavior
            document.querySelectorAll('.error-item[data-field]').forEach(item => {
                item.style.cursor = 'pointer';
                item.addEventListener('click', function () {
                    const fieldId = this.getAttribute('data-field');
                    const targetModalId = this.getAttribute('data-modal');
                    const fieldElement = document.getElementById(fieldId);

                    if (fieldElement) {
                        errorValidationModal.hide(); // Hide error modal

                        // If the target is in another modal, ensure it's shown
                        if (targetModalId && !document.getElementById(targetModalId).classList.contains('show')) {
                            bootstrap.Modal.getInstance(document.getElementById(targetModalId)).show();
                        }

                        setTimeout(() => {
                            fieldElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
                            fieldElement.focus();
                            fieldElement.classList.add('pulse-error');
                            setTimeout(() => fieldElement.classList.remove('pulse-error'), 1000);
                        }, 500); // Delay to allow modal to show
                    }
                });
            });
        }

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
        document.addEventListener('click', function(event) {
            const sidebar = document.getElementById('sidebar');
            const menuToggle = document.querySelector('.menu-toggle');
            
            if (!sidebar.contains(event.target) && !menuToggle.contains(event.target)) {
                if (sidebar.classList.contains('active')) {
                    toggleSidebar();
                }
            }
        });

        document.addEventListener('DOMContentLoaded', initial);
    </script>
</body>
</html>