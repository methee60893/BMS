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
                                    <option selected>USD</option>
                                    <option>THB</option>
                                    <option>EUR</option>
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
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                <button type="button" id="btnConfirmExtra" class="btn btn-primary">Confirm</button>
            </div>
        </div>
    </div>


    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script>
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
                    url: 'Handler/POUploadHandler.ashx?action=preview',
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

        let currencyCal = function () {
            let result = 0.00;
            let amtCCY = !isNaN(txtAmtCCY.value) ? txtAmtCCY.value : "";
            let exRate = txtExRate.value;
            exRate = (exRate == 0) ? 1 : parseFloat(exRate);
            txtAmtTHB.value = amtCCY * exRate;
        }

        let initial = function () {
            const firstMenuLink = document.querySelector('.menu-link');
            if (firstMenuLink) {
                firstMenuLink.classList.add('expanded');
            }


            //InitData master
            InitMSData();
            segmentDropdown.addEventListener('change', changeVendor);
            txtAmtCCY.addEventListener('change', currencyCal);
            txtExRate.addEventListener('change', currencyCal);

            btnSubmit.addEventListener('click', function () {
                // Implement submission logic here

            });

            // *** ADDED: Approve Button Click Event ***
            //btnApprove.addEventListener('click', approveSelectedItems);
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
