<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="createOTBswitching.aspx.vb" Inherits="BMS.createOTBswitching" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BMS - Create OTB Switching</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <style>
        :root {
            --primary-blue: #0B56A4;
            --sidebar-bg: #2c3e50;
            --sidebar-hover: #34495e;
            --orange-header: #D2691E;
            --light-blue-bg: #E6F2FF;
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
.tab-pane { 
    display: none;
}

    .tab-pane.active {
        display: block;
    }

        /* Switch Container */
        .switch-container {
            background: white;
            border-radius: 0 0 8px 8px;
            padding: 30px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }

        .switch-section {
            border: 2px solid #dee2e6;
            border-radius: 8px;
            padding: 25px;
            margin-bottom: 20px;
        }

        .section-title {
            font-size: 1.1rem;
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

        .amount-input {
            background: var(--light-blue-bg);
            border: 2px solid #b8d4f1;
            font-weight: 600;
            color: #2c3e50;
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

        /* Responsive */
        @media (max-width: 768px) {
            .content-area {
                padding: 15px;
            }

            .switch-section {
                padding: 15px;
            }

            .tab-button {
                padding: 10px 15px;
                font-size: 0.9rem;
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
                    <li><a href="createOTBswitching.aspx" class="menu-link active">Create OTB Switching</a></li>
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
                Create OTB Switching
            </div>

            <!-- Tabs -->
            <div class="custom-tabs">
                <button class="tab-button custom-tab active" onclick="switchTab('switching')">
                    <i class="bi bi-arrow-repeat"></i>OTB Switching
                </button>
                <button class="tab-button custom-tab" onclick="switchTab('extra')">
                    <i class="bi bi-plus-circle"></i>Extra Budget
                </button>
            </div>
            <div class="tab-content-area">
                <!-- Switch Tab Content -->
                <div class="tab-pane active" id="switchingTab">
                    <div class="switch-container">
                        <div class="row">
                            <div class="col-12">
                                <!-- From Section -->
                                <div class="switch-section">
                                    <div class="section-title">
                                        <i class="bi bi-box-arrow-right"></i>From
                                    </div>

                                    <div class="row g-3 mb-3">
                                        <div class="col-md-3">
                                            <label class="form-label">Year</label>
                                            <select id="DDYearFrom" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-3">
                                            <label class="form-label">Month</label>
                                            <select id="DDMonthFrom" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-3">
                                            <label class="form-label">Company</label>
                                            <select id="DDCompanyFrom" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-3">
                                        </div>
                                    </div>

                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6">
                                            <label class="form-label">Category</label>
                                            <select id="DDCategoryFrom" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-6">
                                            <label class="form-label">Segment</label>
                                            <select id="DDSegmentFrom" class="form-select">
                                            </select>
                                        </div>
                                    </div>

                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6">
                                            <label class="form-label">Brand</label>
                                            <select id="DDBrandFrom" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-6">
                                            <label class="form-label">Vendor</label>
                                            <select id="DDVendorFrom" class="form-select">
                                            </select>
                                        </div>
                                    </div>
                                </div>

                                <!-- To Section -->
                                <div class="switch-section">
                                    <div class="section-title">
                                        <i class="bi bi-box-arrow-in-right"></i>To
                                    </div>

                                    <div class="row g-3 mb-3">
                                        <div class="col-md-3">
                                            <label class="form-label">Year</label>
                                            <select id="DDYearTo" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-3">
                                            <label class="form-label">Month</label>
                                            <select id="DDMonthTo" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-3">
                                            <label class="form-label">Company</label>
                                            <select id="DDCompanyTo" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-3">
                                        </div>
                                    </div>

                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6">
                                            <label class="form-label">Category</label>
                                            <select id="DDCategoryTo" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-6">
                                            <label class="form-label">Segment</label>
                                            <select id="DDSegmentTo" class="form-select">
                                            </select>
                                        </div>
                                    </div>

                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6">
                                            <label class="form-label">Brand</label>
                                            <select id="DDBrandTo" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-6">
                                            <label class="form-label">Vendor</label>
                                            <select id="DDVendorTo" class="form-select">
                                            </select>
                                        </div>
                                    </div>

                                    <div class="row g-3">
                                        <div class="col-md-3">
                                            <label class="form-label">Amount (THB)</label>
                                            <input id="txtAmontSwitch" type="text" class="form-control amount-input" value="0.00">
                                        </div>
                                    </div>
                                </div>

                                <!-- Submit Button -->
                                <div class="text-end mt-4">
                                    <button type="button" class="btn-submit" id="btnSubmitSwitch">
                                        <i class="bi bi-check-circle"></i>Submit
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Extra Tab Content -->
                <div class="tab-pane" id="extraTab">
                    <div class="switch-container">
                        <div class="row">
                            <div class="col-12">
                                <div class="switch-section">
                                    <div class="section-title">
                                        <i class="bi bi-plus-circle"></i>Extra
                                    </div>

                                    <div class="row g-3 mb-3">
                                        <div class="col-md-3">
                                            <label class="form-label">Year</label>
                                            <select id="DDYearEx" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-3">
                                            <label class="form-label">Month</label>
                                            <select id="DDMonthEx" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-3">
                                            <label class="form-label">Company</label>
                                            <select id="DDCompanyEx" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-3">
                                        </div>
                                    </div>

                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6">
                                            <label class="form-label">Category</label>
                                            <select id="DDCategoryEx" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-6">
                                            <label class="form-label">Segment</label>
                                            <select id="DDSegmentEx" class="form-select">
                                            </select>
                                        </div>
                                    </div>

                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6">
                                            <label class="form-label">Brand</label>
                                            <select id="DDBrandEx" class="form-select">
                                            </select>
                                        </div>
                                        <div class="col-md-6">
                                            <label class="form-label">Vendor</label>
                                            <select id="DDVendorEx" class="form-select">
                                            </select>
                                        </div>
                                    </div>
                                    <div class="row g-3">
                                        <div class="col-md-3">
                                            <label class="form-label">Amount (THB)</label>
                                            <input id="txtAmontEx" type="text" class="form-control amount-input" value="0.00">
                                        </div>
                                    </div>
                                </div>

                                <!-- Submit Button -->
                                <div class="text-end mt-4">
                                    <button type="button" class="btn-submit" id="btnSubmitExtra">
                                        <i class="bi bi-check-circle"></i>Submit
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
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
    <div class="modal fade" id="previewSwitchModal" tabindex="-1" aria-labelledby="previewSwitchModalLabel" data-bs-backdrop="static" data-bs-keyboard="false" >
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="previewSwitchModalLabel">Preview Data</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div id="previewSwitchContainer">
                        <div class="col-12">
                            <!-- From Section -->
                            <div class="switch-section">
                                <div class="section-title">
                                    <i class="bi bi-box-arrow-right"></i>From
                                </div>

                                <div class="row g-3 mb-3">
                                    <div class="col-md-3">
                                        <label class="form-label">Year</label>
                                        <input id="tsYearFrom" type="text" class="form-control" value="">
                                    </div>
                                    <div class="col-md-3">
                                        <label class="form-label">Month</label>
                                        <input id="tsMonthFrom" type="text" class="form-control " value="">
                                    </div>
                                    <div class="col-md-3">
                                        <label class="form-label">Company</label>
                                        <input id="tsCompanyFrom" type="text" class="form-control" value="">
                                    </div>
                                    <div class="col-md-3">
                                    </div>
                                </div>

                                <div class="row g-3 mb-3">
                                    <div class="col-md-6">
                                        <label class="form-label">Category</label>
                                        <input id="tsCategoryFrom" type="text" class="form-control " value="">
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">Segment</label>
                                        <input id="tsSegmentFrom" type="text" class="form-control " value="">
                                    </div>
                                </div>

                                <div class="row g-3 mb-3">
                                    <div class="col-md-6">
                                        <label class="form-label">Brand</label>
                                        <input id="tsBrandFrom" type="text" class="form-control" value="">
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">Vendor</label>
                                        <input id="tsVendorFrom" type="text" class="form-control " value="">
                                    </div>
                                </div>
                            </div>

                            <!-- To Section -->
                            <div class="switch-section">
                                <div class="section-title">
                                    <i class="bi bi-box-arrow-in-right"></i>To
                                </div>

                                <div class="row g-3 mb-3">
                                    <div class="col-md-3">
                                        <label class="form-label">Year</label>
                                        <input id="tsYearTo" type="text" class="form-control" value="">
                                    </div>
                                    <div class="col-md-3">
                                        <label class="form-label">Month</label>
                                        <input id="tsMonthTo" type="text" class="form-control " value="">
                                    </div>
                                    <div class="col-md-3">
                                        <label class="form-label">Company</label>
                                        <input id="tsCompanyTo" type="text" class="form-control " value="">
                                    </div>
                                    <div class="col-md-3">
                                    </div>
                                </div>

                                <div class="row g-3 mb-3">
                                    <div class="col-md-6">
                                        <label class="form-label">Category</label>
                                        <input id="tsCategoryTo" type="text" class="form-control" value="">
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">Segment</label>
                                        <input id="tsSegmentTo" type="text" class="form-control" value="">
                                    </div>
                                </div>

                                <div class="row g-3 mb-3">
                                    <div class="col-md-6">
                                        <label class="form-label">Brand</label>
                                        <input id="tsBrandTo" type="text" class="form-control " value="">
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">Vendor</label>
                                        <input id="tsVendorTo" type="text" class="form-control" value="">
                                    </div>
                                </div>

                                <div class="row g-3">
                                    <div class="col-md-3">
                                        <label class="form-label">Amount (THB)</label>
                                        <input id="tsAmontSwitch" type="text" class="form-control " value="0.00">
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="button" id="btnConfirmSwitch" class="btn btn-primary">Confirm</button>
                </div>
            </div>
        </div>
    </div>
    <div class="modal fade" id="previewExtraModal" tabindex="-1" aria-labelledby="previewModalLabel" data-bs-backdrop="static" data-bs-keyboard="false" >
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="previewExtraModalLabel">Preview Data</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div id="previewExtraContainer">
                        <div class="col-12">
                            <div class="switch-section">
                                <div class="section-title">
                                    <i class="bi bi-plus-circle"></i>Extra
                                </div>

                                <div class="row g-3 mb-3">
                                    <div class="col-md-3">
                                        <label class="form-label">Year</label>
                                        <input id="tsYearEx" type="text" class="form-control" value="">
                                    </div>
                                    <div class="col-md-3">
                                        <label class="form-label">Month</label>
                                        <input id="tsMonthEx" type="text" class="form-control" value="">
                                    </div>
                                    <div class="col-md-3">
                                        <label class="form-label">Company</label>
                                        <input id="tsCompanyEx" type="text" class="form-control" value="">
                                    </div>
                                    <div class="col-md-3">
                                    </div>
                                </div>

                                <div class="row g-3 mb-3">
                                    <div class="col-md-6">
                                        <label class="form-label">Category</label>
                                        <input id="tsCategoryEx" type="text" class="form-control" value="">
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">Segment</label>
                                        <input id="tsSegmentEx" type="text" class="form-control" value="">
                                    </div>
                                </div>

                                <div class="row g-3 mb-3">
                                    <div class="col-md-6">
                                        <label class="form-label">Brand</label>
                                        <input id="tsBrandEx" type="text" class="form-control" value="">
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">Vendor</label>
                                        <input id="tsVendorEx" type="text" class="form-control" value="">
                                    </div>
                                </div>
                                <div class="row g-3">
                                    <div class="col-md-3">
                                        <label class="form-label">Amount (THB)</label>
                                        <input id="tsAmontEx" type="text" class="form-control" value="0.00">
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
    </div>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // ==========================================
        // Global Variables
        // ==========================================
        var yearDropdownf, monthDropdownf, companyDropdownf, categoryDropdownf, segmentDropdownf, brandDropdownf, vendorDropdownf;
        var yearDropdownt, monthDropdownt, companyDropdownt, categoryDropdownt, segmentDropdownt, brandDropdownt, vendorDropdownt;
        var yearDropdownE, monthDropdownE, companyDropdownE, categoryDropdownE, segmentDropdownE, brandDropdownE, vendorDropdownE;
        var txtAmontSwitch, txtAmontEx;
        var btnSubmit, btnSubmitExtra;

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

        // ==========================================
        // Tab Switching Function
        // ==========================================
        function switchTab(tabName) {
            // Hide all tab panes
            var tabPanes = document.querySelectorAll('.tab-pane');
            for (var i = 0; i < tabPanes.length; i++) {
                tabPanes[i].classList.remove('active');
            }

            // Remove active class from all tab buttons
            var tabButtons = document.querySelectorAll('.custom-tab');
            for (var j = 0; j < tabButtons.length; j++) {
                tabButtons[j].classList.remove('active');
            }

            // Show selected tab by index
            if (tabName === 'switching') {
                if (tabPanes[0]) tabPanes[0].classList.add('active');
                if (tabButtons[0]) tabButtons[0].classList.add('active');
            } else if (tabName === 'extra') {
                if (tabPanes[1]) tabPanes[1].classList.add('active');
                if (tabButtons[1]) tabButtons[1].classList.add('active');
            }
        }

        // ==========================================
        // Initialize Function
        // ==========================================
        var initial = function () {
            // From Section
            yearDropdownf = document.getElementById("DDYearFrom");
            monthDropdownf = document.getElementById("DDMonthFrom");
            companyDropdownf = document.getElementById("DDCompanyFrom");
            categoryDropdownf = document.getElementById("DDCategoryFrom");
            segmentDropdownf = document.getElementById("DDSegmentFrom");
            brandDropdownf = document.getElementById("DDBrandFrom");
            vendorDropdownf = document.getElementById("DDVendorFrom");
            txtAmontSwitch = document.getElementById("txtAmontSwitch");
            btnSubmit = document.getElementById("btnSubmitSwitch");

            // To Section
            yearDropdownt = document.getElementById("DDYearTo");
            monthDropdownt = document.getElementById("DDMonthTo");
            companyDropdownt = document.getElementById("DDCompanyTo");
            categoryDropdownt = document.getElementById("DDCategoryTo");
            segmentDropdownt = document.getElementById("DDSegmentTo");
            brandDropdownt = document.getElementById("DDBrandTo");
            vendorDropdownt = document.getElementById("DDVendorTo");

            // Extra Section
            yearDropdownE = document.getElementById("DDYearEx");
            monthDropdownE = document.getElementById("DDMonthEx");
            companyDropdownE = document.getElementById("DDCompanyEx");
            categoryDropdownE = document.getElementById("DDCategoryEx");
            segmentDropdownE = document.getElementById("DDSegmentEx");
            brandDropdownE = document.getElementById("DDBrandEx");
            vendorDropdownE = document.getElementById("DDVendorEx");
            txtAmontEx = document.getElementById("txtAmontEx");
            btnSubmitExtra = document.getElementById("btnSubmitExtra");

            // Init Master Data
            InitMSData();

            // Event Listeners
            if (segmentDropdownf) segmentDropdownf.addEventListener('change', changeVendorF);
            if (segmentDropdownt) segmentDropdownt.addEventListener('change', changeVendorT);
            if (segmentDropdownE) segmentDropdownE.addEventListener('change', changeVendorE);

            // Submit Buttons
            if (btnSubmit) {
                btnSubmit.addEventListener('click', handleSwitchSubmit);
            }

            if (btnSubmitExtra) {
                btnSubmitExtra.addEventListener('click', handleExtraSubmit);
            }

            // Event Listeners for Modal Confirm Buttons
            var btnConfirmSwitch = document.getElementById('btnConfirmSwitch');
            if (btnConfirmSwitch) {
                btnConfirmSwitch.addEventListener('click', saveSwitchingData);
            }

            var btnConfirmExtra = document.getElementById('btnConfirmExtra');
            if (btnConfirmExtra) {
                btnConfirmExtra.addEventListener('click', saveExtraData);
            }
        };

        // ==========================================
        // Handle Switch Submit
        // ==========================================
        async function handleSwitchSubmit(e) {
            e.preventDefault();
            clearValidationErrors();

            var formData = new FormData();
            formData.append('yearFrom', yearDropdownf.value);
            formData.append('monthFrom', monthDropdownf.value);
            formData.append('companyFrom', companyDropdownf.value);
            formData.append('categoryFrom', categoryDropdownf.value);
            formData.append('segmentFrom', segmentDropdownf.value);
            formData.append('brandFrom', brandDropdownf.value);
            formData.append('vendorFrom', vendorDropdownf.value);
            formData.append('yearTo', yearDropdownt.value);
            formData.append('monthTo', monthDropdownt.value);
            formData.append('companyTo', companyDropdownt.value);
            formData.append('categoryTo', categoryDropdownt.value);
            formData.append('segmentTo', segmentDropdownt.value);
            formData.append('brandTo', brandDropdownt.value);
            formData.append('vendorTo', vendorDropdownt.value);
            formData.append('amount', txtAmontSwitch.value);

            try {
                showLoading(true);
                var response = await fetch('Handler/ValidateHandler.ashx?action=validateSwitch', {
                    method: 'POST',
                    body: formData
                });
                var result = await response.json();
                showLoading(false);

                if (result.success) {
                    populatePreviewData();
                    var previewSwitchModal = new bootstrap.Modal(document.getElementById('previewSwitchModal'), {
                        keyboard: false
                    });
                    previewSwitchModal.show();
                } else {
                    showErrorModal(result.errors, 'Switch Transaction', '', result.availableBudget);
                }
            } catch (error) {
                showLoading(false);
                console.error('Validation error:', error);
                showErrorModal({ 'general': 'Failed to validate data: ' + error.message }, 'System Error');
            }
        }

        // ==========================================
        // Handle Extra Submit
        // ==========================================
        async function handleExtraSubmit(e) {
            e.preventDefault();
            clearValidationErrors();

            var formData = new FormData();
            formData.append('year', yearDropdownE.value);
            formData.append('month', monthDropdownE.value);
            formData.append('company', companyDropdownE.value);
            formData.append('category', categoryDropdownE.value);
            formData.append('segment', segmentDropdownE.value);
            formData.append('brand', brandDropdownE.value);
            formData.append('vendor', vendorDropdownE.value);
            formData.append('amount', txtAmontEx.value);

            try {
                showLoading(true);
                var response = await fetch('Handler/ValidateHandler.ashx?action=validateExtra', {
                    method: 'POST',
                    body: formData
                });
                var result = await response.json();
                showLoading(false);

                if (result.success) {
                    populateExtraPreviewData();
                    var previewExtraModal = new bootstrap.Modal(document.getElementById('previewExtraModal'), {
                        keyboard: false
                    });
                    previewExtraModal.show();
                } else {
                    showErrorModal(result.errors, 'Extra Budget', 'Ex', result.currentBudget);
                }
            } catch (error) {
                showLoading(false);
                console.error('Validation error:', error);
                showErrorModal({ 'general': 'Failed to validate data: ' + error.message }, 'System Error');
            }
        }

        // ==========================================
        // Show Error Modal
        // ==========================================
        function showErrorModal(errors, transactionType, suffix, availableBudget) {
            transactionType = transactionType || 'Transaction';
            suffix = suffix || '';

            var fieldInfo = {
                'yearFrom': { name: 'Year', section: 'From', element: 'DDYear' },
                'monthFrom': { name: 'Month', section: 'From', element: 'DDMonth' },
                'companyFrom': { name: 'Company', section: 'From', element: 'DDCompany' },
                'categoryFrom': { name: 'Category', section: 'From', element: 'DDCategory' },
                'segmentFrom': { name: 'Segment', section: 'From', element: 'DDSegment' },
                'brandFrom': { name: 'Brand', section: 'From', element: 'DDBrand' },
                'vendorFrom': { name: 'Vendor', section: 'From', element: 'DDVendor' },
                'yearTo': { name: 'Year', section: 'To', element: 'DDYeart' },
                'monthTo': { name: 'Month', section: 'To', element: 'DDMontht' },
                'companyTo': { name: 'Company', section: 'To', element: 'DDCompanyt' },
                'categoryTo': { name: 'Category', section: 'To', element: 'DDCategoryt' },
                'segmentTo': { name: 'Segment', section: 'To', element: 'DDSegmentt' },
                'brandTo': { name: 'Brand', section: 'To', element: 'DDBrandt' },
                'vendorTo': { name: 'Vendor', section: 'To', element: 'DDVendort' },
                'amount': { name: 'Amount (THB)', section: 'General', element: 'txtAmontSwitch' },
                'year': { name: 'Year', section: 'Extra', element: 'DDYearEx' },
                'month': { name: 'Month', section: 'Extra', element: 'DDMonthEx' },
                'company': { name: 'Company', section: 'Extra', element: 'DDCompanyEx' },
                'category': { name: 'Category', section: 'Extra', element: 'DDCategoryEx' },
                'segment': { name: 'Segment', section: 'Extra', element: 'DDSegmentEx' },
                'brand': { name: 'Brand', section: 'Extra', element: 'DDBrandEx' },
                'vendor': { name: 'Vendor', section: 'Extra', element: 'DDVendorEx' }
            };

            var errorCount = Object.keys(errors).length;
            var summaryText = document.getElementById('errorSummaryText');
            if (summaryText) {
                var summary = 'Found ' + errorCount + ' validation error' + (errorCount > 1 ? 's' : '') + ' in your ' + transactionType;
                if (availableBudget !== null && availableBudget !== undefined) {
                    var budgetFormatted = parseFloat(availableBudget).toLocaleString('en-US', {
                        minimumFractionDigits: 2,
                        maximumFractionDigits: 2
                    });
                    summary += '<br><small class="text-muted">Available Budget: <strong>' + budgetFormatted + ' THB</strong></small>';
                }
                summaryText.innerHTML = summary;
            }

            var errorListContainer = document.getElementById('errorListContainer');
            var errorHtml = '';

            var sortedErrors = Object.entries(errors).sort(function (a, b) {
                var sectionA = fieldInfo[a[0]] ? fieldInfo[a[0]].section : 'General';
                var sectionB = fieldInfo[b[0]] ? fieldInfo[b[0]].section : 'General';
                var sectionOrder = { 'From': 1, 'To': 2, 'General': 3, 'Extra': 4 };
                return (sectionOrder[sectionA] || 99) - (sectionOrder[sectionB] || 99);
            });

            for (var i = 0; i < sortedErrors.length; i++) {
                var field = sortedErrors[i][0];
                var message = sortedErrors[i][1];
                var info = fieldInfo[field];

                if (field === 'general') {
                    errorHtml += '<div class="error-item">' +
                        '<div class="error-item-icon"><i class="bi bi-exclamation-circle"></i></div>' +
                        '<div class="error-item-content">' +
                        '<div class="error-field-name"><span class="error-section-badge">General</span>Validation Error</div>' +
                        '<p class="error-message">' + message + '</p>' +
                        '</div></div>';
                } else if (info) {
                    var sectionColor = info.section === 'From' ? '#FF6B35' : info.section === 'To' ? '#4ECDC4' : '#dc3545';
                    errorHtml += '<div class="error-item" data-field="' + info.element + '">' +
                        '<div class="error-item-icon"><i class="bi bi-x-circle"></i></div>' +
                        '<div class="error-item-content">' +
                        '<div class="error-field-name">' +
                        '<span class="error-section-badge" style="background-color: ' + sectionColor + ';">' +
                        info.section + '</span>' + info.name + '</div>' +
                        '<p class="error-message">' + message + '</p>' +
                        '</div></div>';

                    var element = document.getElementById(info.element);
                    if (element) element.classList.add('has-error');
                }
            }

            if (errorListContainer) errorListContainer.innerHTML = errorHtml;

            var errorModal = new bootstrap.Modal(document.getElementById('errorValidationModal'), {
                backdrop: 'static',
                keyboard: false
            });
            errorModal.show();

            var errorItems = document.querySelectorAll('.error-item[data-field]');
            for (var j = 0; j < errorItems.length; j++) {
                (function (item) {
                    item.style.cursor = 'pointer';
                    item.addEventListener('click', function () {
                        var fieldId = this.getAttribute('data-field');
                        var fieldElement = document.getElementById(fieldId);
                        if (fieldElement) {
                            errorModal.hide();
                            setTimeout(function () {
                                fieldElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
                                fieldElement.focus();
                                fieldElement.classList.add('pulse-error');
                                setTimeout(function () {
                                    fieldElement.classList.remove('pulse-error');
                                }, 1000);
                            }, 300);
                        }
                    });
                })(errorItems[j]);
            }
        }

        // ==========================================
        // Helper Functions
        // ==========================================
        function clearValidationErrors() {
            var hasErrorElements = document.querySelectorAll('.has-error');
            for (var i = 0; i < hasErrorElements.length; i++) {
                hasErrorElements[i].classList.remove('has-error');
            }
            var pulseErrorElements = document.querySelectorAll('.pulse-error');
            for (var j = 0; j < pulseErrorElements.length; j++) {
                pulseErrorElements[j].classList.remove('pulse-error');
            }
        }

        function populatePreviewData() {
            document.getElementById("tsYearFrom").value = yearDropdownf.value;
            document.getElementById("tsMonthFrom").value = monthDropdownf.value;
            document.getElementById("tsCompanyFrom").value = companyDropdownf.value;
            document.getElementById("tsCategoryFrom").value = categoryDropdownf.value;
            document.getElementById("tsSegmentFrom").value = segmentDropdownf.value;
            document.getElementById("tsBrandFrom").value = brandDropdownf.value;
            document.getElementById("tsVendorFrom").value = vendorDropdownf.value;
            document.getElementById("tsYearTo").value = yearDropdownt.value;
            document.getElementById("tsMonthTo").value = monthDropdownt.value;
            document.getElementById("tsCompanyTo").value = companyDropdownt.value;
            document.getElementById("tsCategoryTo").value = categoryDropdownt.value;
            document.getElementById("tsSegmentTo").value = segmentDropdownt.value;
            document.getElementById("tsBrandTo").value = brandDropdownt.value;
            document.getElementById("tsVendorTo").value = vendorDropdownt.value;
            document.getElementById("tsAmontSwitch").value = txtAmontSwitch.value;
        }

        function populateExtraPreviewData() {
            document.getElementById("tsYearEx").value = yearDropdownE.value;
            document.getElementById("tsMonthEx").value = monthDropdownE.value;
            document.getElementById("tsCompanyEx").value = companyDropdownE.value;
            document.getElementById("tsCategoryEx").value = categoryDropdownE.value;
            document.getElementById("tsSegmentEx").value = segmentDropdownE.value;
            document.getElementById("tsBrandEx").value = brandDropdownE.value;
            document.getElementById("tsVendorEx").value = vendorDropdownE.value;
            document.getElementById("tsAmontEx").value = txtAmontEx.value;
        }

        function showLoading(show) {
            var loadingHtml = '<div id="loadingOverlay" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 9999; display: flex; flex-direction: column; align-items: center; justify-content: center;">' +
                '<div class="spinner-border text-light" role="status" style="width: 3rem; height: 3rem;">' +
                '<span class="visually-hidden">Loading...</span></div>' +
                '<p class="text-light mt-3 mb-0">Validating data...</p></div>';

            if (show) {
                document.body.insertAdjacentHTML('beforeend', loadingHtml);
            } else {
                var overlay = document.getElementById('loadingOverlay');
                if (overlay) overlay.remove();
            }
        }

        // ==========================================
        // Handle Save Data (After Confirm)
        // ==========================================
        async function saveSwitchingData() {
            showLoading(true);
            var formData = new FormData();

            // ดึงข้อมูลจาก hidden fields ใน Modal
            formData.append('yearFrom', document.getElementById('tsYearFrom').value);
            formData.append('monthFrom', document.getElementById('tsMonthFrom').value);
            formData.append('companyFrom', document.getElementById('tsCompanyFrom').value);
            formData.append('categoryFrom', document.getElementById('tsCategoryFrom').value);
            formData.append('segmentFrom', document.getElementById('tsSegmentFrom').value);
            formData.append('brandFrom', document.getElementById('tsBrandFrom').value);
            formData.append('vendorFrom', document.getElementById('tsVendorFrom').value);

            formData.append('yearTo', document.getElementById('tsYearTo').value);
            formData.append('monthTo', document.getElementById('tsMonthTo').value);
            formData.append('companyTo', document.getElementById('tsCompanyTo').value);
            formData.append('categoryTo', document.getElementById('tsCategoryTo').value);
            formData.append('segmentTo', document.getElementById('tsSegmentTo').value);
            formData.append('brandTo', document.getElementById('tsBrandTo').value);
            formData.append('vendorTo', document.getElementById('tsVendorTo').value);

            formData.append('amount', document.getElementById('tsAmontSwitch').value);
            formData.append('createdBy', 'System'); // TODO: เปลี่ยนเป็น User จริง
            formData.append('remark', ''); // TODO: เพิ่มช่อง Remark ถ้าต้องการ

            try {
                var response = await fetch('Handler/SaveOTBHandler.ashx?action=saveSwitching', {
                    method: 'POST',
                    body: formData
                });
                var result = await response.json();
                showLoading(false);

                if (result.success) {
                    bootstrap.Modal.getInstance(document.getElementById('previewSwitchModal')).hide();
                    alert(result.message || 'Save successful!');

                    // ล้างฟอร์ม
                    // ล้างฟอร์ม (Manual Reset)
                    // Clear "From" fields
                    if (yearDropdownf) yearDropdownf.value = "";
                    if (monthDropdownf) monthDropdownf.value = "";
                    if (companyDropdownf) companyDropdownf.value = "";
                    if (categoryDropdownf) categoryDropdownf.value = "";
                    if (segmentDropdownf) segmentDropdownf.value = "";
                    if (brandDropdownf) brandDropdownf.value = "";
                    if (vendorDropdownf) vendorDropdownf.value = "";

                    // Clear "To" fields
                    if (yearDropdownt) yearDropdownt.value = "";
                    if (monthDropdownt) monthDropdownt.value = "";
                    if (companyDropdownt) companyDropdownt.value = "";
                    if (categoryDropdownt) categoryDropdownt.value = "";
                    if (segmentDropdownt) segmentDropdownt.value = "";
                    if (brandDropdownt) brandDropdownt.value = "";
                    if (vendorDropdownt) vendorDropdownt.value = "";

                    // Clear Amount (ใช้ ID จาก HTML ที่คุณส่งมาให้)
                    var txtAmontSwitch = document.getElementById("txtAmontSwitch");
                    if (txtAmontSwitch) txtAmontSwitch.value = "0.00";
                    // --- ✨ จบส่วนที่แก้ไข ---

                    InitMSData(); // โหลด Master data ใหม่
                } else {
                    alert('Save failed: ' + result.message);
                }
            } catch (error) {
                showLoading(false);
                alert('Error: ' + error.message);
            }
        }

        async function saveExtraData() {
            showLoading(true);
            var formData = new FormData();

            // ดึงข้อมูลจาก hidden fields ใน Modal
            formData.append('year', document.getElementById('tsYearEx').value);
            formData.append('month', document.getElementById('tsMonthEx').value);
            formData.append('company', document.getElementById('tsCompanyEx').value);
            formData.append('category', document.getElementById('tsCategoryEx').value);
            formData.append('segment', document.getElementById('tsSegmentEx').value);
            formData.append('brand', document.getElementById('tsBrandEx').value);
            formData.append('vendor', document.getElementById('tsVendorEx').value);
            formData.append('amount', document.getElementById('tsAmontEx').value);

            formData.append('createdBy', 'System'); // TODO: เปลี่ยนเป็น User จริง
            formData.append('remark', ''); // TODO: เพิ่มช่อง Remark ถ้าต้องการ

            try {
                var response = await fetch('Handler/SaveOTBHandler.ashx?action=saveExtra', {
                    method: 'POST',
                    body: formData
                });
                var result = await response.json();
                showLoading(false);

                if (result.success) {
                    bootstrap.Modal.getInstance(document.getElementById('previewExtraModal')).hide();
                    alert(result.message || 'Extra budget save successful!');
                    // ล้างฟอร์ม
                    // ล้างฟอร์ม (Manual Reset)
                    if (yearDropdownE) yearDropdownE.value = "";
                    if (monthDropdownE) monthDropdownE.value = "";
                    if (companyDropdownE) companyDropdownE.value = "";
                    if (categoryDropdownE) categoryDropdownE.value = "";
                    if (segmentDropdownE) segmentDropdownE.value = "";
                    if (brandDropdownE) brandDropdownE.value = "";
                    if (vendorDropdownE) vendorDropdownE.value = "";

                    // Clear Amount (ใช้ ID จาก HTML ที่คุณส่งมาให้)
                    var txtAmontEx = document.getElementById("txtAmontEx");
                    if (txtAmontEx) txtAmontEx.value = "0.00";
                    InitMSData(); // โหลด Master data ใหม่
                } else {
                    alert('Save failed: ' + result.message);
                }
            } catch (error) {
                showLoading(false);
                alert('Error: ' + error.message);
            }
        }

        // ==========================================
        // Initialize Master Data
        // ==========================================
        var InitMSData = function () {
            InitSegment(segmentDropdownf);
            InitCategoty(categoryDropdownf);
            InitBrand(brandDropdownf);
            InitVendor(vendorDropdownf);
            InitMSYear(yearDropdownf);
            InitMonth(monthDropdownf);
            InitCompany(companyDropdownf);

            InitSegment(segmentDropdownt);
            InitCategoty(categoryDropdownt);
            InitBrand(brandDropdownt);
            InitVendor(vendorDropdownt);
            InitMSYear(yearDropdownt);
            InitMonth(monthDropdownt);
            InitCompany(companyDropdownt);

            InitSegment(segmentDropdownE);
            InitCategoty(categoryDropdownE);
            InitBrand(brandDropdownE);
            InitVendor(vendorDropdownE);
            InitMSYear(yearDropdownE);
            InitMonth(monthDropdownE);
            InitCompany(companyDropdownE);
        };

        // Master Data Functions
        var InitSegment = function (segmentddf) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=SegmentMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    segmentddf.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        };

        var InitMSYear = function (yearddf) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=YearMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    yearddf.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        };

        var InitMonth = function (monthddf) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=MonthMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    monthddf.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        };

        var InitCompany = function (companyddf) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=CompanyMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    companyddf.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        };

        var InitCategoty = function (categoryddf) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=CategoryMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    categoryddf.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        };

        var InitBrand = function (branddd) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=BrandMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    branddd.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        };

        var InitVendor = function (vendorddf) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=VendorMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    vendorddf.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        };

        // Change Vendor Functions
        var changeVendorF = function () {
            var segmentCode = segmentDropdownf.value;
            if (!segmentCode) {
                InitVendor(vendorDropdownf);
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
                    vendorDropdownf.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        };

        var changeVendorT = function () {
            var segmentCode = segmentDropdownt.value;
            if (!segmentCode) {
                InitVendor(vendorDropdownt);
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
                    vendorDropdownt.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        };

        var changeVendorE = function () {
            var segmentCode = segmentDropdownE.value;
            if (!segmentCode) {
                InitVendor(vendorDropdownE);
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
                    vendorDropdownE.innerHTML = response;
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        };

        // Document Ready
        document.addEventListener('DOMContentLoaded', initial);
    </script>
</body>
</html>
