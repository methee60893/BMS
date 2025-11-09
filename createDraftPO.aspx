<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="createDraftPO.aspx.vb" Inherits="BMS.createDraftPO" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BMS - Create Draft PO</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <style>
        :root {
            --primary-blue: #0B56A4;
            --sidebar-bg: #2c3e50;
            --sidebar-hover: #34495e;
            --orange-header: #D2691E;
            --light-blue-bg: #E6F2FF;
            --green-btn: #28a745;
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

        /* Tabs */
        .custom-tabs {
            display: flex;
            background: white;
            border-bottom: 2px solid #dee2e6;
            margin-bottom: 0;
        }

        .tab-button {
            padding: 12px 30px;
            background: #6c757d;
            color: white;
            border: none;
            cursor: pointer;
            font-weight: 600;
            transition: all 0.3s;
            flex: 1;
        }

            .tab-button.active {
                background: var(--primary-blue);
            }

            .tab-button:hover {
                background: var(--primary-blue);
                opacity: 0.9;
            }

        /* Tab Content */
        .tab-content {
            display: none;
        }

            .tab-content.active {
                display: block;
            }

        /* Form Container */
        .form-container {
            background: white;
            border-radius: 0 0 8px 8px;
            padding: 30px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }

        .form-section {
            border: 2px solid #dee2e6;
            border-radius: 8px;
            padding: 25px;
        }

        .section-title {
            font-size: 1rem;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid var(--primary-blue);
        }

        /* Form Styles */
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

        /* Form Row */
        .form-row-item {
            margin-bottom: 15px;
        }

        .form-row-display {
            display: flex;
            gap: 20px;
            align-items: start;
        }

            .form-row-display .form-group {
                display: flex;
                gap: 15px;
                align-items: center;
                flex: 1;
            }

                .form-row-display .form-group label {
                    min-width: 100px;
                    margin-bottom: 0;
                    font-weight: 600;
                }

                .form-row-display .form-group .info-display {
                    flex: 1;
                }

                .form-row-display .form-group input,
                .form-row-display .form-group select {
                    flex: 1;
                }

        /* Buttons */
        .btn-submit {
            background: var(--primary-blue);
            color: white;
            padding: 12px 40px;
            border: none;
            border-radius: 6px;
            font-weight: 600;
            font-size: 1rem;
            transition: all 0.3s;
            cursor: pointer;
        }

            .btn-submit:hover {
                background: #094580;
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(11,86,164,0.3);
            }

        .btn-upload {
            background: var(--green-btn);
            color: white;
            padding: 10px 30px;
            border: none;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.95rem;
            transition: all 0.3s;
            cursor: pointer;
        }

            .btn-upload:hover {
                background: #218838;
                transform: translateY(-2px);
            }

        /* Upload Section */
        .upload-section {
            padding: 40px 20px;
        }

        .upload-title {
            font-size: 1.1rem;
            font-weight: 600;
            color: #2c3e50;
            margin-bottom: 25px;
        }

        .file-input-group {
            display: flex;
            align-items: center;
            gap: 15px;
        }

            .file-input-group label {
                min-width: 80px;
                font-weight: 600;
                margin-bottom: 0;
            }

            .file-input-group input[type="file"] {
                flex: 1;
            }

        /* Responsive */
        @media (max-width: 768px) {
            .content-area {
                padding: 15px;
            }

            .form-section {
                padding: 15px;
            }

            .tab-button {
                padding: 10px 15px;
                font-size: 0.9rem;
            }

            .form-row-display {
                flex-direction: column;
                gap: 15px;
            }

            .file-input-group {
                flex-direction: column;
                align-items: stretch;
            }

                .file-input-group label {
                    min-width: auto;
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

        .table-responsive { border: 1px solid #dee2e6; }
        .sticky-header { position: sticky; top: 0; z-index: 10; }
        .text-truncate-custom { max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    </style>
</head>
<body>
    <!-- Sidebar Overlay -->
    <div class="sidebar-overlay" id="sidebarOverlay" onclick="toggleSidebar()"></div>

    <!-- Sidebar -->
    <div class="sidebar" id="sidebar">
        <div class="sidebar-header">
            <h3><i class="bi bi-building"></i>BMS</h3>
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
                    <li><a href="createDraftPO.aspx" class="menu-link active">Create Draft PO</a></li>
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
                Create Draft PO
            </div>

            <!-- Tabs -->
            <div class="custom-tabs">
                <button class="tab-button active" onclick="switchTab('txn')">
                    <i class="bi bi-pencil-square"></i>Create by TXN
                </button>
                <button class="tab-button" onclick="switchTab('upload')">
                    <i class="bi bi-cloud-upload"></i>Upload file
                </button>
            </div>

            <!-- Create by TXN Tab Content -->
            <div class="tab-content active" id="txnTab">
                <div class="form-container">
                    <form id="poTxnForm">
                        <div class="form-section">
                            <div class="section-title">
                                Create Draft PO by transaction
                            </div>

                            <!-- Row 1 -->
                            <div class="form-row-display form-row-item">
                                <div class="form-group">
                                    <label>Year</label>
                                    <select id="DDYear" class="form-select">
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label>Category</label>
                                    <select id="DDCategory" class="form-select">
                                    </select>
                                </div>
                            </div>

                            <!-- Row 2 -->
                            <div class="form-row-display form-row-item">
                                <div class="form-group">
                                    <label>Month</label>
                                    <select id="DDMonth" class="form-select">
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label>Segment</label>
                                    <select id="DDSegment" class="form-select">
                                    </select>
                                </div>
                            </div>

                            <!-- Row 3 -->
                            <div class="form-row-display form-row-item">
                                <div class="form-group">
                                    <label>Company</label>
                                    <select id="DDCompany" class="form-select">
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label>Brand</label>
                                    <select id="DDBrand" class="form-select">
                                    </select>
                                </div>
                            </div>

                            <!-- Row 4 -->
                            <div class="form-row-display form-row-item">
                                <div class="form-group" style="visibility: hidden;">
                                    <label>-</label>
                                    <div class="info-display">-</div>
                                </div>
                                <div class="form-group">
                                    <label>Vendor</label>
                                    <select id="DDVendor" class="form-select">
                                    </select>
                                </div>
                            </div>

                            <!-- Row 5 -->
                            <div class="form-row-display form-row-item">
                                <div class="form-group">
                                    <label>Draft PO no.</label>
                                    <input id="txtPONO" type="text" class="form-control" placeholder="Enter PO number">
                                </div>
                                <div class="form-group">
                                    <label>Amount (CCY)</label>
                                    <input id="txtAmtCCY" type="text" class="form-control" pattern="^\d+(\.\d{1,2})?$" title="Enter a valid amount (e.g., 123 or 123.45)" placeholder="0.00">
                                </div>
                            </div>

                            <!-- Row 6 -->
                            <div class="form-row-display form-row-item">
                                <div class="form-group">
                                    <label>CCY</label>
                                    <select id="DDCCY" class="form-select">
                                        
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label>Amount (THB)</label>
                                    <input id="txtAmtTHB" type="text" class="form-control" placeholder="0.00" readonly>
                                </div>
                            </div>

                            <!-- Row 7 -->
                            <div class="form-row-display form-row-item">
                                <div class="form-group">
                                    <label>Exchange rate</label>
                                    <input id="txtExRate" type="text" class="form-control" placeholder="0.00" pattern="^\d+(\.\d{1,2})?$" title="Enter a valid amount (e.g., 123 or 123.45)">
                                </div>
                                <div class="form-group">
                                    <label>Remark</label>
                                    <input id="txtRemark" type="text" class="form-control" placeholder="Enter remark">
                                </div>
                            </div>

                            <!-- Submit Button -->
                            <div class="text-center mt-4">
                                <button type="button" class="btn-submit" id="btnSubmit">
                                    <i class="bi bi-check-circle"></i>Submit
                                </button>
                            </div>
                        </div>
                    </form>
                    
                </div>
            </div>

            <!-- Upload File Tab Content -->
            <div class="tab-content" id="uploadTab">
                <div class="form-container">
                    <div class="form-section">
                        <div class="upload-section">
                            <div class="upload-title">
                                Create Draft PO by upload
                            </div>

                            <div class="file-input-group">
                                <label>File</label>
                                <input type="file" id="fileUpload" class="form-control" accept=".xlsx,.xls,.csv">
                                <button type="button" id="btnUpload" class="btn-upload">
                                    <i class="bi bi-upload"></i>Upload
                                </button>
                            </div>
                        </div>
                    </div>
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

    <div class="modal fade" id="previewPOTXNModal" tabindex="-1" aria-labelledby="previewPOTXNModalLabel" data-bs-backdrop="static" data-bs-keyboard="false">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="previewPOTXNModalLabel">Preview Data</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div id="previewPOTXNContainer">
                        <div class="form-container">
                            <div class="form-section">
                                <div class="section-title">
                                    Create Draft PO by transaction
                                </div>

                                <!-- Row 1 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label>Year</label>
                                        <input id="tsYear" type="text" class="form-control" readonly>
                                    </div>
                                    <div class="form-group">
                                        <label>Category</label>
                                        <input id="tsCategory" type="text" class="form-control" readonly>
                                    </div>
                                </div>

                                <!-- Row 2 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label>Month</label>
                                        <input id="tsMonth" type="text" class="form-control" readonly>
                                    </div>
                                    <div class="form-group">
                                        <label>Segment</label>
                                        <input id="tsSegment" type="text" class="form-control" readonly>
                                    </div>
                                </div>

                                <!-- Row 3 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label>Company</label>
                                        <input id="tsCompany" type="text" class="form-control" readonly>
                                    </div>
                                    <div class="form-group">
                                        <label>Brand</label>
                                        <input id="tsBrand" type="text" class="form-control" readonly>
                                    </div>
                                </div>

                                <!-- Row 4 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group" style="visibility: hidden;">
                                        <label>-</label>
                                        <div class="info-display">-</div>
                                    </div>
                                    <div class="form-group">
                                        <label>Vendor</label>
                                        <input id="tsVendor" type="text" class="form-control" readonly>
                                    </div>
                                </div>

                                <!-- Row 5 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label>Draft PO no.</label>
                                        <input id="tsPONO" type="text" class="form-control" placeholder="Enter PO number">
                                    </div>
                                    <div class="form-group">
                                        <label>Amount (CCY)</label>
                                        <input id="tsAmtCCY" type="text" class="form-control" placeholder="0.00">
                                    </div>
                                </div>

                                <!-- Row 6 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label>CCY</label>
                                        <input id="tsCCY" type="text" class="form-control" placeholder="Enter PO number">
                                    </div>
                                    <div class="form-group">
                                        <label>Amount (THB)</label>
                                        <input id="tsAmtTHB" type="text" class="form-control" placeholder="0.00" readonly>
                                    </div>
                                </div>

                                <!-- Row 7 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label>Exchange rate</label>
                                        <input id="tsExRate" type="text" class="form-control" placeholder="0.00" pattern="^\d+(\.\d{1,2})?$" title="Enter a valid amount (e.g., 123 or 123.45)">
                                    </div>
                                    <div class="form-group">
                                        <label>Remark</label>
                                        <input id="tsRemark" type="text" class="form-control" placeholder="Enter remark">
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="button" id="btnConfirm" class="btn btn-primary">Confirm</button>
                </div>
            </div>
            
        </div>
    </div>

    <!-- Loading Overlay -->
    <div class="loading-overlay" id="loadingOverlay">
        <div class="loading-content">
            <div class="loading-spinner"></div>
            <p class="loading-text" id="loadingText">Processing...</p>
            <p class="loading-subtext" id="loadingSubtext">Please wait a moment</p>
        </div>
    </div>

     <!-- Error Validation Modal -->
    <div class="modal fade" id="errorValidationModal" tabindex="-1" data-bs-backdrop="static" data-bs-keyboard="false">
        <div class="modal-dialog modal-dialog-centered modal-lg">
            <div class="modal-content">
                <div class="modal-header bg-danger text-white">
                    <h5 class="modal-title">
                        <i class="bi bi-exclamation-triangle-fill me-2"></i>
                        Validation Error
                    </h5>
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
    
    <!-- Success Modal -->
    <div class="modal fade" id="successModal" tabindex="-1" data-bs-backdrop="static" data-bs-keyboard="false">
        <div class="modal-dialog modal-dialog-centered">
            <div class="modal-content">
                <div class="modal-header bg-success text-white">
                    <h5 class="modal-title" id="successModalTitle">
                        <i class="bi bi-check-circle-fill me-2"></i>
                        Success
                    </h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <p id="successModalMessage">Operation completed successfully.</p>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-success" data-bs-dismiss="modal">
                        <i class="bi bi-hand-thumbs-up me-2"></i>OK
                    </button>
                </div>
            </div>
        </div>
    </div>


    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>

        // ADDED: Modal instances
        let previewTxnModal, errorModal, successModal, uploadPreviewModal;
        let previewPOTXNModal;



        let yearDropdown = document.getElementById("DDYear");
        let monthDropdown = document.getElementById("DDMonth");
        let companyDropdown = document.getElementById("DDCompany");
        let segmentDropdown = document.getElementById("DDSegment");
        let categoryDropdown = document.getElementById("DDCategory");
        let brandDropdown = document.getElementById("DDBrand");
        let vendorDropdown = document.getElementById("DDVendor");
        let ccyDropdown = document.getElementById("DDCCY");
        let txtPONO = document.getElementById("txtPONO");
        let txtAmtCCY = document.getElementById("txtAmtCCY");
        let txtAmtTHB = document.getElementById("txtAmtTHB");
        let txtExRate = document.getElementById("txtExRate");
        let txtRemark = document.getElementById("txtRemark");

        let btnSubmitData = document.getElementById('btnSubmitData');
        let btnSubmit = document.getElementById('btnSubmit');
        let btnUpload = document.getElementById('btnUpload');
        let fileUploadInput = document.getElementById('fileUpload');
        let previewTableContainer = document.getElementById('previewTableContainer');


        function handleUploadPreview() {
            var fileInput = fileUploadInput;
            var file = fileInput.files[0];
            var currentUser = '<%= If(Session("user") IsNot Nothing, HttpUtility.JavaScriptStringEncode(Session("user").ToString()), "unknown") %>';
                    var uploadBy = currentUser || 'unknown';

                    if (!file) {
                        showErrorModal({ 'general': 'Please select a file to upload.' }, 'Upload Error');
                        return;
                    }

                    showLoading(true, 'Uploading file...');

                    var formData = new FormData();
                    formData.append('file', file);
                    formData.append('uploadBy', uploadBy);

                    $.ajax({
                        url: 'Handler/POUploadHandler.ashx?action=preview',
                        type: 'POST',
                        data: formData,
                        processData: false,
                        contentType: false,
                        success: function (response) {
                            showLoading(false);
                            previewTableContainer.innerHTML = response;
                            uploadPreviewModal.show();
                        },
                        error: function (xhr, status, error) {
                            showLoading(false);
                            previewTableContainer.innerHTML = '';
                            console.error('Error loading preview:', error);
                            showErrorModal({ 'general': 'Error loading preview: ' + xhr.responseText }, 'Upload Error');
                        }
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

        // --- Loading Overlay Function ---
        function showLoading(show = true, message = 'Processing...', subMessage = 'Please wait') {
            const overlay = document.getElementById('loadingOverlay');
            if (overlay) {
                document.getElementById('loadingText').textContent = message;
                document.getElementById('loadingSubtext').textContent = subMessage;
                if (show) {
                    overlay.classList.add('active');
                } else {
                    overlay.classList.remove('active');
                }
            }
        }

        // --- Error Modal Function ---
        function showErrorModal(errors, transactionType) {
            let errorCount = Object.keys(errors).length;
            let summaryText = document.getElementById('errorSummaryText');
            summaryText.textContent = `Found ${errorCount} error(s) in your ${transactionType}.`;

            let errorListContainer = document.getElementById('errorListContainer');
            errorListContainer.innerHTML = ''; // Clear previous errors

            // Define field mappings
            const fieldMap = {
                'DDYear': 'Year',
                'DDCategory': 'Category',
                'DDMonth': 'Month',
                'DDSegment': 'Segment',
                'DDCompany': 'Company',
                'DDBrand': 'Brand',
                'DDVendor': 'Vendor',
                'txtPONO': 'Draft PO no.',
                'txtAmtCCY': 'Amount (CCY)',
                'DDCCY': 'CCY',
                'txtExRate': 'Exchange rate',
                'general': 'General Error'
            };

            for (const fieldId in errors) {
                const fieldName = fieldMap[fieldId] || fieldId;
                const message = errors[fieldId];

                const errorItem = document.createElement('div');
                errorItem.className = 'error-item';
                errorItem.setAttribute('data-field-id', fieldId);
                errorItem.innerHTML = `
                    <div class="error-item-icon"><i class="bi bi-x-circle"></i></div>
                    <div class="error-item-content">
                        <div class="error-field-name">${fieldName}</div>
                        <p class="error-message">${message}</p>
                    </div>
                `;
                errorListContainer.appendChild(errorItem);

                // Highlight the field
                const fieldElement = document.getElementById(fieldId);
                if (fieldElement) {
                    fieldElement.classList.add('has-error');
                }
            }
            errorModal.show();
        }

        function showErrorSaveModal(title, transactionType) {
           

            let errorListContainer = document.getElementById('errorListContainer');
            errorListContainer.innerHTML = ''; // Clear previous errors

                const errorItem = document.createElement('div');
                errorItem.className = 'error-item';
                errorItem.innerHTML = `
             <div class="error-item-icon"><i class="bi bi-x-circle"></i></div>
             <div class="error-item-content">
                 <div class="error-field-name">${transactionType}</div>
                 <p class="error-message">${title}</p>
             </div>
         `;
                errorListContainer.appendChild(errorItem);

                //// Highlight the field
                //const fieldElement = document.getElementById(fieldId);
                //if (fieldElement) {
                //    fieldElement.classList.add('has-error');
                //}
            
            errorModal.show();
        }

        // --- Success Modal Function ---
        function showSuccessModal(title, message) {
            document.getElementById('successModalTitle').textContent = title;
            document.getElementById('successModalMessage').textContent = message;
            successModal.show();
        }

        // --- Field Error Clear Function ---
        function clearValidationErrors() {
            document.querySelectorAll('.has-error').forEach(el => el.classList.remove('has-error'));
        }

        let currencyCal = function () {
            let result = 0.00;
            let amtCCY = parseFloat(txtAmtCCY.value) || 0;
            let exRate = parseFloat(txtExRate.value) || 0;

            // Auto-set ExRate to 1 if CCY is THB
            if (ccyDropdown.value === 'THB') {
                txtExRate.value = "1.00";
                exRate = 1.00;
                txtExRate.readOnly = true;
            } else {
                txtExRate.readOnly = false;
                // if exRate was 1, clear it so user can input
                if (exRate === 1.00 && txtExRate.value === "1.00") {
                    //txtExRate.value = ""; 
                }
            }

            txtAmtTHB.value = (amtCCY * exRate).toFixed(2);
        }

        let initial = function () {
            const firstMenuLink = document.querySelector('.menu-link');
            if (firstMenuLink) {
                firstMenuLink.classList.add('expanded');
            }


            //InitData master
            InitMSData();

            previewPOTXNModal = new bootstrap.Modal(document.getElementById('previewPOTXNModal'));
            errorModal = new bootstrap.Modal(document.getElementById('errorValidationModal'));
            successModal = new bootstrap.Modal(document.getElementById('successModal'));

            uploadPreviewModal = new bootstrap.Modal(document.getElementById('previewModal'));

            segmentDropdown.addEventListener('change', changeVendor);
            vendorDropdown.addEventListener('change', chengeCCY);
            txtAmtCCY.addEventListener('change', currencyCal);
            txtExRate.addEventListener('change', currencyCal);
            ccyDropdown.addEventListener('change', currencyCal);
            btnSubmit.addEventListener('click', handleSubmitPOTXN);

            // ADDED: Confirm button logic
            document.getElementById('btnConfirm').addEventListener('click', handleConfirmSavePOTXN);
            btnSubmitData.addEventListener('click', handleUploadSubmit);
            btnUpload.addEventListener('click', handleUploadPreview);
        }

        let InitMSData = function () {
            InitSegment(segmentDropdown);
            InitCategoty(categoryDropdown);
            InitBrand(brandDropdown);
            InitVendor(vendorDropdown);
            InitMSYear(yearDropdown);
            InitMonth(monthDropdown);
            InitCompany(companyDropdown);
            InitCCY(ccyDropdown);
        }

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
            // Implement month initialization if needed
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

        let InitCCY = function (ccyDropdown) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=CCYMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    ccyDropdown.innerHTML = response;
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

        let chengeCCY = function () {
           var vendorCode = vendorDropdown.value;
            if (!vendorCode) {
                // ถ้าไม่มีค่า ให้โหลด vendor ทั้งหมด
                InitVendor(vendorDropdown);
                return;
            }
            var formData = new FormData();
            formData.append('vendorCode', vendorCode);
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=CCYMSListChg',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    ccyDropdown.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }

       

        let InsertPreview = function () {
            var formData = new FormData();
            formData.append('segmentCode', "");
            $.ajax({
                url: 'Handler/DataPOHandler.ashx?action=InsertPreview',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }

        let PoPreview = function () {
            var formData = new FormData();
            formData.append('segmentCode', "");
            $.ajax({
                url: 'Handler/DataPOHandler.ashx?action=PoPreview',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {

                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }

        // ==========================================
        // ===== NEW: PO TXN SUBMIT FUNCTIONS =====
        // ==========================================

        /**
         * Step 1: Handle the "Submit" button click.
         * Validates data via handler, then shows preview or error modal.
         */
        async function handleSubmitPOTXN(e) {
            e.preventDefault();
            clearValidationErrors();
            showLoading(true, "Validating...");

            const formData = new FormData();
            formData.append('year', yearDropdown.value);
            formData.append('month', monthDropdown.value);
            formData.append('company', companyDropdown.value);
            formData.append('category', categoryDropdown.value);
            formData.append('segment', segmentDropdown.value);
            formData.append('brand', brandDropdown.value);
            formData.append('vendor', vendorDropdown.value);
            formData.append('pono', txtPONO.value);
            formData.append('amtCCY', txtAmtCCY.value);
            formData.append('ccy', ccyDropdown.value);
            formData.append('exRate', txtExRate.value);
            formData.append('remark', txtRemark.value);

            try {
                const response = await fetch('Handler/ValidateHandler.ashx?action=validateDraftPO', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }

                const result = await response.json();
                showLoading(false);

                if (result.success) {
                    populatePOTXNPreview();
                    previewPOTXNModal.show();
                } else {
                    showErrorModal(result.errors, 'System Error');
                    showValidationErrorsOnForm(result.errors);
                }
            } catch (error) {
                showLoading(false);
                console.error('Validation fetch error:', error);
                showErrorModal({ 'general': 'A system error occurred: ' + error.message }, 'System Error');
            }
        }

        /**
         * Step 2: Populate the preview modal with form data.
         */
        function populatePOTXNPreview() {
            // Helper to get selected text from a dropdown
            const getSelectedText = (el) => el.options[el.selectedIndex]?.text || el.value;

            document.getElementById("tsYear").value = getSelectedText(yearDropdown);
            document.getElementById("tsMonth").value = getSelectedText(monthDropdown);
            document.getElementById("tsCompany").value = getSelectedText(companyDropdown);
            document.getElementById("tsCategory").value = getSelectedText(categoryDropdown);
            document.getElementById("tsSegment").value = getSelectedText(segmentDropdown);
            document.getElementById("tsBrand").value = getSelectedText(brandDropdown);
            document.getElementById("tsVendor").value = getSelectedText(vendorDropdown);
            document.getElementById("tsPONO").value = txtPONO.value;
            document.getElementById("tsAmtCCY").value = parseFloat(txtAmtCCY.value || 0).toFixed(2);
            document.getElementById("tsCCY").value = getSelectedText(ccyDropdown);
            document.getElementById("tsExRate").value = parseFloat(txtExRate.value || 0).toFixed(4);
            document.getElementById("tsAmtTHB").value = txtAmtTHB.value;
            document.getElementById("tsRemark").value = txtRemark.value;
        }

        /**
         * Step 3: Handle the "Confirm & Save" button click from the preview modal.
         * Saves data via handler, then shows success or error.
         */
        async function handleConfirmSavePOTXN() {
            showLoading(true, "Saving...");

            // Data is sent from the *preview modal* inputs to ensure it matches what user confirmed
            const formData = new FormData();
            formData.append('year', yearDropdown.value); // Send value, not text
            formData.append('month', monthDropdown.value);
            formData.append('company', companyDropdown.value);
            formData.append('category', categoryDropdown.value);
            formData.append('segment', segmentDropdown.value);
            formData.append('brand', brandDropdown.value);
            formData.append('vendor', vendorDropdown.value);
            formData.append('pono', document.getElementById("tsPONO").value);
            formData.append('amtCCY', document.getElementById("tsAmtCCY").value);
            formData.append('ccy', ccyDropdown.value); // Send value
            formData.append('exRate', document.getElementById("tsExRate").value);
            formData.append('amtTHB', document.getElementById("tsAmtTHB").value);
            formData.append('remark', document.getElementById("tsRemark").value);
            // createdBy will be handled by the server using Session

            try {
                const response = await fetch('Handler/DataPOHandler.ashx?action=saveDraftPOTXN', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }

                const result = await response.json();
                showLoading(false);
                previewPOTXNModal.hide();

                if (result.success) {
                    showSuccessModal('Success', result.message);
                    clearPOForm();
                } else {
                    // Show save error in the main error modal
                    showErrorModal({ 'general': result.message }, 'Save Error');
                }

            } catch (error) {
                showLoading(false);
                previewPOTXNModal.hide();
                console.error('Save error:', error);
                showErrorModal({ 'general': 'Failed to save data: ' + error.message }, 'System Error');
            }
        }

        /**
         * Clears the main PO form after successful save.
         */
        function clearPOForm() {
            document.getElementById('poTxnForm').reset();
            // Reset calculated fields
            txtAmtTHB.value = "0.00";
            // Manually trigger change on dropdowns to reset dependent fields if necessary
            segmentDropdown.dispatchEvent(new Event('change'));
            ccyDropdown.dispatchEvent(new Event('change'));
            // Clear any lingering validation
            clearValidationErrors();
        }

        function clearValidationErrors() {
            document.querySelectorAll('.has-error').forEach(el => el.classList.remove('has-error'));
            document.querySelectorAll('.validation-message').forEach(el => {
                el.textContent = '';
                el.style.display = 'none';
            });
        }

        function showValidationErrorsOnForm(errors) {
            for (const field in errors) {
                const el = document.querySelector(`.form-control[id="txt${field.toUpperCase()}"], .form-select[id="DD${field.charAt(0).toUpperCase() + field.slice(1)}"]`);
                const msgEl = document.querySelector(`.validation-message[data-field="${field}"]`);

                if (el) {
                    el.classList.add('has-error');
                }
                if (msgEl) {
                    msgEl.textContent = errors[field];
                    msgEl.style.display = 'block';
                }
            }
        }

        function handleUploadSubmit(e) {
            e.preventDefault();

            var selectedRowsData = [];
            // Find checked checkboxes
            $('#previewTableContainer input[name="selectedRows"]:checked').each(function () {
                var cb = $(this);
                // Get data from data attributes
                var rowData = {
                    DraftPONo: cb.data('pono'),
                    Year: cb.data('year'),
                    Month: cb.data('month'),
                    Category: cb.data('category'),
                    Company: cb.data('company'),
                    Segment: cb.data('segment'),
                    Brand: cb.data('brand'),
                    Vendor: cb.data('vendor'),
                    AmountTHB: cb.data('amountthb'),
                    AmountCCY: cb.data('amountccy'),
                    CCY: cb.data('ccy'),
                    ExRate: cb.data('exrate'),
                    Remark: cb.data('remark')
                };
                selectedRowsData.push(rowData);
            });

            if (selectedRowsData.length === 0) {

                showErrorSaveModal('Please select at least one valid row to save.', 'warning');
                return;
            }

            // ** FIX for Session("user") NullReferenceException **
            var uploadBy = '<%= If(Session("user") IsNot Nothing, HttpUtility.JavaScriptStringEncode(Session("user").ToString()), "unknown") %>';

            var formData = new FormData();
            formData.append('uploadBy', uploadBy);
            formData.append('selectedData', JSON.stringify(selectedRowsData)); // Send data as JSON string

            showLoading(true, "Saving...", "Submitting selected rows...");
            btnSubmitData.disabled = true;

            $.ajax({
                url: 'Handler/POUploadHandler.ashx?action=savePreview',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    showLoading(false);
                    btnSubmitData.disabled = false;
                    uploadPreviewModal.hide();
                    previewTableContainer.innerHTML = '';
                    fileUploadInput.value = ''; // Clear file input
                    
                    showSuccessModal(response, 'success') 
                },
                error: function (xhr, status, error) {
                    showLoading(false);
                    btnSubmitData.disabled = false;
                    // Display error in a more persistent way
                    showErrorSaveModal(`Error saving data: ${xhr.responseText || error}`, 'danger');
                }
            });
        }
        // --- END: Upload File Logic ---

       


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
