<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="createOTBswitching.aspx.vb" Inherits="BMS.createOTBswitching" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - Create OTB Switching</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="style/theme.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
</head>
<body>
    <div class="sidebar-overlay" id="sidebarOverlay" onclick="toggleSidebar()"></div>

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
                    <li id="menuDraftOTBPlan" runat="server" ><a href="draftOTB.aspx" class="menu-link">Draft OTB Plan</a></li>
                    <li id="menuApprovedOTBPlan" runat="server" ><a href="approvedOTB.aspx" class="menu-link">Approved OTB Plan</a></li>
                </ul>
            </li>
            <li class="menu-item">
                <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'otbSwitching')">
                    <i class="bi bi-arrow-left-right"></i>
                    <span>OTB Switching</span>
                    <i class="bi bi-chevron-down"></i>
                </a>
                <ul class="submenu" id="otbSwitching">
                    <li id="menuCreateOTBSwitching" runat="server" ><a href="createOTBswitching.aspx" class="menu-link active">Create OTB Switching</a></li>
                    <li id="menuSwitchingTransaction" runat="server" ><a href="transactionOTBSwitching.aspx" class="menu-link">Switching Transaction</a></li>
                </ul>
            </li>
            <li class="menu-item" id="grpmenuPO" runat="server" >
                <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'po')">
                    <i class="bi bi-file-earmark-text"></i>
                    <span>PO</span>
                    <i class="bi bi-chevron-down"></i>
                </a>
                <ul class="submenu" id="po">
                    <li id="menuCreateDraftPO" runat="server"><a href="createDraftPO.aspx" class="menu-link">Create Draft PO</a></li>
                    <li id="menuDraftPO" runat="server" ><a href="draftPO.aspx" class="menu-link">Draft PO</a></li>
                    <li id="menuMatchActualPO" runat="server" ><a href="matchActualPO.aspx" class="menu-link">Match Actual PO</a></li>
                    <li id="menuActualPO" runat="server" ><a href="actualPO.aspx" class="menu-link">Actual PO</a></li>
                </ul>
            </li>
            <li class="menu-item">
                <a href="otbRemaining.aspx" class="menu-link">
                    <i class="bi bi-bar-chart-line"></i>
                    <span>OTB Remaining</span>
                </a>
            </li>
            <li class="menu-item" id="grpmenuMaster" runat="server">
                <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'master')">
                    <i class="bi bi-database"></i>
                    <span>Master File</span>
                    <i class="bi bi-chevron-down"></i>
                </a>
                <ul class="submenu" id="master">
                    <li id="menuVendor" runat="server" ><a href="master_vendor.aspx" class="menu-link">Master Vendor</a></li>
                    <li id="menuBrand" runat="server" ><a href="master_brand.aspx" class="menu-link">Master Brand</a></li>
                    <li id="menuCategory" runat="server" ><a href="master_category.aspx" class="menu-link">Master Category</a></li>
                </ul>
            </li>
            <li class="menu-item"><a href="default.aspx" class="menu-link"><i class="bi bi-box-arrow-left"></i> Logout</a></li>
        </ul>
    </div>

    <div class="main-wrapper">
        <div class="top-navbar">
            <div class="d-flex align-items-center gap-3">
                <button class="menu-toggle" onclick="toggleSidebar()">
                    <i class="bi bi-list"></i>
                </button>
                <h1 class="page-title" id="pageTitle">KBMS - Create OTB Switching</h1>
            </div>
            <div class="user-info">
                <span class="d-none d-md-inline">Welcome, <%= HttpUtility.JavaScriptStringEncode(Session("fullname").ToString()) %></span>
                <div class="user-avatar">
                    <i class="bi bi-person-circle"></i>
                </div>
            </div>
        </div>

        <div class="content-area">
            <div class="page-header">
                Create OTB Switching
            </div>

            <div class="custom-tabs">
                <button class="tab-button active" onclick="switchTab('switching')">
                    <i class="bi bi-arrow-repeat"></i> OTB Switching
                </button>
                <button class="tab-button" onclick="switchTab('extra')">
                    <i class="bi bi-plus-circle"></i> Extra Budget
                </button>
                <button class="tab-button" onclick="switchTab('bulk')">
                    <i class="bi bi-file-earmark-spreadsheet"></i> Bulk Upload
                </button>
            </div>

            <div class="tab-content-area">
                <div class="tab-content active" id="switchingTab">
                    <div class="form-container">
                        <div class="row">
                            <div class="col-12">
                                <div class="switch-section">
                                    <div class="section-title">
                                        <i class="bi bi-box-arrow-right"></i> Out
                                    </div>
                                    <div class="row g-3 mb-3">
                                        <div class="col-md-3"><label class="form-label">Year</label><select id="DDYearFrom" class="form-select"></select></div>
                                        <div class="col-md-3"><label class="form-label">Month</label><select id="DDMonthFrom" class="form-select"></select></div>
                                        <div class="col-md-3"><label class="form-label">Company</label><select id="DDCompanyFrom" class="form-select"></select></div>
                                        <div class="col-md-3"></div>
                                    </div>
                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6"><label class="form-label">Category</label><select id="DDCategoryFrom" class="form-select"></select></div>
                                        <div class="col-md-6"><label class="form-label">Segment</label><select id="DDSegmentFrom" class="form-select"></select></div>
                                    </div>
                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6"><label class="form-label">Brand</label><select id="DDBrandFrom" class="form-select"></select></div>
                                        <div class="col-md-6"><label class="form-label">Vendor</label><select id="DDVendorFrom" class="form-select"></select></div>
                                    </div>
                                </div>

                                <div class="switch-section">
                                    <div class="section-title">
                                        <i class="bi bi-box-arrow-in-right"></i> In
                                    </div>
                                    <div class="row g-3 mb-3">
                                        <div class="col-md-3"><label class="form-label">Year</label><select id="DDYearTo" class="form-select"></select></div>
                                        <div class="col-md-3"><label class="form-label">Month</label><select id="DDMonthTo" class="form-select"></select></div>
                                        <div class="col-md-3"><label class="form-label">Company</label><select id="DDCompanyTo" class="form-select"></select></div>
                                        <div class="col-md-3"></div>
                                    </div>
                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6"><label class="form-label">Category</label><select id="DDCategoryTo" class="form-select"></select></div>
                                        <div class="col-md-6"><label class="form-label">Segment</label><select id="DDSegmentTo" class="form-select"></select></div>
                                    </div>
                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6"><label class="form-label">Brand</label><select id="DDBrandTo" class="form-select"></select></div>
                                        <div class="col-md-6"><label class="form-label">Vendor</label><select id="DDVendorTo" class="form-select"></select></div>
                                    </div>
                                    <div class="row g-3">
                                        <div class="col-md-3"><label class="form-label">Amount (THB)</label><input id="txtAmontSwitch" type="text" class="form-control amount-input" value="0.00" autocomplete="off"></div>
                                    </div>
                                </div>

                                <div class="text-end mt-4">
                                    <button type="button" class="btn-submit" id="btnSubmitSwitch">
                                        <i class="bi bi-check-circle"></i> Submit
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="tab-content" id="extraTab">
                    <div class="form-container">
                        <div class="row">
                            <div class="col-12">
                                <div class="switch-section">
                                    <div class="section-title">
                                        <i class="bi bi-plus-circle"></i> Extra
                                    </div>
                                    <div class="row g-3 mb-3">
                                        <div class="col-md-3"><label class="form-label">Year</label><select id="DDYearEx" class="form-select"></select></div>
                                        <div class="col-md-3"><label class="form-label">Month</label><select id="DDMonthEx" class="form-select"></select></div>
                                        <div class="col-md-3"><label class="form-label">Company</label><select id="DDCompanyEx" class="form-select"></select></div>
                                        <div class="col-md-3"></div>
                                    </div>
                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6"><label class="form-label">Category</label><select id="DDCategoryEx" class="form-select"></select></div>
                                        <div class="col-md-6"><label class="form-label">Segment</label><select id="DDSegmentEx" class="form-select"></select></div>
                                    </div>
                                    <div class="row g-3 mb-3">
                                        <div class="col-md-6"><label class="form-label">Brand</label><select id="DDBrandEx" class="form-select"></select></div>
                                        <div class="col-md-6"><label class="form-label">Vendor</label><select id="DDVendorEx" class="form-select"></select></div>
                                    </div>
                                    <div class="row g-3">
                                        <div class="col-md-3"><label class="form-label">Amount (THB)</label><input id="txtAmontEx" type="text" class="form-control amount-input" value="0.00" autocomplete="off"></div>
                                    </div>
                                </div>
                                <div class="text-end mt-4">
                                    <button type="button" class="btn-submit" id="btnSubmitExtra">
                                        <i class="bi bi-check-circle"></i> Submit
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="tab-content" id="bulkTab">
                    <div class="form-container">
                        <div class="row">
                            <div class="col-12">
                                <div class="mb-3">
                                    <label class="form-label fw-bold">Upload Excel File (Switch / Extra)</label>
                                    <input type="file" class="form-control" id="fileBulkSwitch" accept=".xlsx, .xls">
                                    <div class="form-text">
                                        Template Columns: <strong>Type</strong> (Switch/Extra), Year, Month, Company, Category, Segment, Brand, Vendor, To_Year, To_Month, ..., Amount, Remark
                                    </div>
                                </div>
                                <div class="text-end">
                                    <button type="button" class="btn btn-info text-white" id="btnPreviewBulk">
                                        <i class="bi bi-eye"></i> Preview
                                    </button>
                                </div>
                            </div>
                        </div>
                        
                        <div id="bulkPreviewContainer" class="mt-4"></div>

                        <div id="bulkResultContainer" class="mt-4" style="display:none;">
                            <h5>Upload Results (SAP)</h5>
                            <div class="table-responsive" style="max-height:400px;">
                                <table class="table table-bordered table-sm" id="tblBulkResult">
                                    <thead class="table-success sticky-top">
                                        <tr>
                                            <th>Row</th>
                                            <th>Type</th>
                                            <th>From</th>
                                            <th>To</th>
                                            <th class="text-end">Amount</th>
                                            <th>Status</th>
                                            <th>Message</th>
                                        </tr>
                                    </thead>
                                    <tbody></tbody>
                                </table>
                            </div>
                        </div>

                        <div class="text-end mt-3" id="divBulkActions" style="display:none;">
                            <button type="button" class="btn-submit" id="btnSaveBulk">
                                <i class="bi bi-cloud-upload"></i> Confirm & Save to SAP
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="errorValidationModal" tabindex="-1" data-bs-backdrop="static" data-bs-keyboard="false">
        <div class="modal-dialog modal-dialog-centered modal-lg">
            <div class="modal-content">
                <div class="modal-header bg-danger text-white">
                    <h5 class="modal-title" id="errorValidationModalTitle">Validation Error</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="alert alert-danger d-flex align-items-center mb-3" role="alert">
                        <i class="bi bi-x-circle-fill fs-4 me-3"></i>
                        <div>
                            <strong id="errorSummaryTitle">Please correct the following errors:</strong>
                            <p class="mb-0 mt-1" id="errorSummaryText">Some fields require your attention</p>
                        </div>
                    </div>
                    <div id="errorListContainer" class="error-list-container"></div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Go Back to Fix</button>
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="previewSwitchModal" tabindex="-1" aria-labelledby="previewSwitchModalLabel" data-bs-backdrop="static" data-bs-keyboard="false" >
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">Preview Data</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div id="previewSwitchContainer">
                        <div class="row"><div class="col-12"><div class="switch-section"><div class="section-title"><i class="bi bi-box-arrow-right"></i>From</div><div class="row g-3 mb-3"><div class="col-md-3"><label class="form-label">Year</label><input id="tsYearFrom" type="text" class="form-control" readonly></div><div class="col-md-3"><label class="form-label">Month</label><input id="tsMonthFrom" type="text" class="form-control" readonly></div><div class="col-md-3"><label class="form-label">Company</label><input id="tsCompanyFrom" type="text" class="form-control" readonly></div><div class="col-md-3"></div></div><div class="row g-3 mb-3"><div class="col-md-6"><label class="form-label">Category</label><input id="tsCategoryFrom" type="text" class="form-control" readonly></div><div class="col-md-6"><label class="form-label">Segment</label><input id="tsSegmentFrom" type="text" class="form-control" readonly></div></div><div class="row g-3 mb-3"><div class="col-md-6"><label class="form-label">Brand</label><input id="tsBrandFrom" type="text" class="form-control" readonly></div><div class="col-md-6"><label class="form-label">Vendor</label><input id="tsVendorFrom" type="text" class="form-control" readonly></div></div></div><div class="switch-section"><div class="section-title"><i class="bi bi-box-arrow-in-right"></i>To</div><div class="row g-3 mb-3"><div class="col-md-3"><label class="form-label">Year</label><input id="tsYearTo" type="text" class="form-control" readonly></div><div class="col-md-3"><label class="form-label">Month</label><input id="tsMonthTo" type="text" class="form-control" readonly></div><div class="col-md-3"><label class="form-label">Company</label><input id="tsCompanyTo" type="text" class="form-control" readonly></div><div class="col-md-3"></div></div><div class="row g-3 mb-3"><div class="col-md-6"><label class="form-label">Category</label><input id="tsCategoryTo" type="text" class="form-control" readonly></div><div class="col-md-6"><label class="form-label">Segment</label><input id="tsSegmentTo" type="text" class="form-control" readonly></div></div><div class="row g-3 mb-3"><div class="col-md-6"><label class="form-label">Brand</label><input id="tsBrandTo" type="text" class="form-control" readonly></div><div class="col-md-6"><label class="form-label">Vendor</label><input id="tsVendorTo" type="text" class="form-control" readonly></div></div><div class="row g-3"><div class="col-md-3"><label class="form-label">Amount (THB)</label><input id="tsAmontSwitch" type="text" class="form-control" readonly></div></div></div></div></div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="button" id="btnConfirmSwitch" class="btn btn-primary">Confirm</button>
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="previewExtraModal" tabindex="-1" data-bs-backdrop="static" data-bs-keyboard="false" >
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">Preview Data</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div id="previewExtraContainer">
                        <div class="row"><div class="col-12"><div class="switch-section"><div class="section-title"><i class="bi bi-plus-circle"></i>Extra</div><div class="row g-3 mb-3"><div class="col-md-3"><label class="form-label">Year</label><input id="tsYearEx" type="text" class="form-control" readonly></div><div class="col-md-3"><label class="form-label">Month</label><input id="tsMonthEx" type="text" class="form-control" readonly></div><div class="col-md-3"><label class="form-label">Company</label><input id="tsCompanyEx" type="text" class="form-control" readonly></div><div class="col-md-3"></div></div><div class="row g-3 mb-3"><div class="col-md-6"><label class="form-label">Category</label><input id="tsCategoryEx" type="text" class="form-control" readonly></div><div class="col-md-6"><label class="form-label">Segment</label><input id="tsSegmentEx" type="text" class="form-control" readonly></div></div><div class="row g-3 mb-3"><div class="col-md-6"><label class="form-label">Brand</label><input id="tsBrandEx" type="text" class="form-control" readonly></div><div class="col-md-6"><label class="form-label">Vendor</label><input id="tsVendorEx" type="text" class="form-control" readonly></div></div><div class="row g-3"><div class="col-md-3"><label class="form-label">Amount (THB)</label><input id="tsAmontEx" type="text" class="form-control" readonly></div></div></div></div></div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="button" id="btnConfirmExtra" class="btn btn-primary">Confirm</button>
                </div>
            </div>
        </div>
    </div>

    <div class="loading-overlay" id="loadingOverlay">
        <div class="loading-content">
            <div class="loading-spinner"></div>
            <p class="loading-text" id="loadingText">Processing...</p>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script>
        // ==========================================
        // Global Variables
        // ==========================================
        var yearDropdownf, monthDropdownf, companyDropdownf, categoryDropdownf, segmentDropdownf, brandDropdownf, vendorDropdownf;
        var yearDropdownt, monthDropdownt, companyDropdownt, categoryDropdownt, segmentDropdownt, brandDropdownt, vendorDropdownt;
        var yearDropdownE, monthDropdownE, companyDropdownE, categoryDropdownE, segmentDropdownE, brandDropdownE, vendorDropdownE;
        var txtAmontSwitch, txtAmontEx;
        var btnSubmit, btnSubmitExtra;
        var errorModal; 

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

        // ==========================================
        // Tab Switching Function
        // ==========================================
        function switchTab(tab) {
            const tabs = document.querySelectorAll('.tab-button');
            tabs.forEach(t => t.classList.remove('active'));

            const tabContents = document.querySelectorAll('.tab-content');
            tabContents.forEach(tc => tc.classList.remove('active'));

            event.target.closest('.tab-button').classList.add('active');

            if (tab === 'switching') {
                document.getElementById('switchingTab').classList.add('active');
            } else if (tab === 'extra') {
                document.getElementById('extraTab').classList.add('active');
            } else if (tab === 'bulk') {
                document.getElementById('bulkTab').classList.add('active');
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

            // Init Error Modal
            errorModal = new bootstrap.Modal(document.getElementById('errorValidationModal'));

            // Currency Listeners
            if (txtAmontSwitch) {
                txtAmontSwitch.addEventListener('keydown', restrictToNumeric);
                txtAmontSwitch.addEventListener('focus', cleanCurrencyOnFocus);
                txtAmontSwitch.addEventListener('blur', formatCurrencyOnBlur);
            }
            if (txtAmontEx) {
                txtAmontEx.addEventListener('keydown', restrictToNumeric);
                txtAmontEx.addEventListener('focus', cleanCurrencyOnFocus);
                txtAmontEx.addEventListener('blur', formatCurrencyOnBlur);
            }

            // Init Select2
            initSelect2([yearDropdownf, yearDropdownt, yearDropdownE, monthDropdownf, monthDropdownt, monthDropdownE,
                         companyDropdownf, companyDropdownt, companyDropdownE, segmentDropdownf, segmentDropdownt, segmentDropdownE,
                         categoryDropdownf, categoryDropdownt, categoryDropdownE, brandDropdownf, brandDropdownt, brandDropdownE,
                         vendorDropdownf, vendorDropdownt, vendorDropdownE]);

            // Init Master Data
            InitMSData();

            // Event Listeners for cascading dropdowns
            if (segmentDropdownf) segmentDropdownf.addEventListener('change', changeVendorF);
            if (segmentDropdownt) segmentDropdownt.addEventListener('change', changeVendorT);
            if (segmentDropdownE) segmentDropdownE.addEventListener('change', changeVendorE);

            // Submit Buttons (Individual)
            if (btnSubmit) btnSubmit.addEventListener('click', handleSwitchSubmit);
            if (btnSubmitExtra) btnSubmitExtra.addEventListener('click', handleExtraSubmit);

            // Modal Confirm Buttons
            var btnConfirmSwitch = document.getElementById('btnConfirmSwitch');
            if (btnConfirmSwitch) btnConfirmSwitch.addEventListener('click', saveSwitchingData);

            var btnConfirmExtra = document.getElementById('btnConfirmExtra');
            if (btnConfirmExtra) btnConfirmExtra.addEventListener('click', saveExtraData);

            // *** NEW BULK UPLOAD LISTENERS ***
            var btnPreviewBulk = document.getElementById('btnPreviewBulk');
            if (btnPreviewBulk) btnPreviewBulk.addEventListener('click', previewBulkFile);

            var btnSaveBulk = document.getElementById('btnSaveBulk');
            if (btnSaveBulk) btnSaveBulk.addEventListener('click', saveBulkData);
        };

        function initSelect2(elements) {
            elements.forEach(el => {
                if (el) $(el).select2({ theme: "bootstrap-5" });
            });
        }

        // ==========================================
        // Bulk Upload Functions (UPDATED FOR JSON)
        // ==========================================
        function previewBulkFile() {
            var fileInput = document.getElementById('fileBulkSwitch');
            if (fileInput.files.length === 0) {
                alert('Please select a file.');
                return;
            }

            var formData = new FormData();
            formData.append('file', fileInput.files[0]);

            showLoading(true, "Reading & Validating File...");
            
            $.ajax({
                url: 'Handler/SwitchUploadHandler.ashx?action=preview',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function(response) {
                    showLoading(false);
                    
                    // 1. ตรวจสอบ response ว่ามาเป็น JSON Object หรือไม่ (jQuery ajax parse ให้แล้วถ้า ContentType ถูก)
                    // ถ้า Server ส่ง string JSON มา อาจต้อง parse เอง แต่ในที่นี้เรากำหนด ContentType = application/json ที่ Server แล้ว
                    
                    if (response.success === false) {
                        alert(response.message);
                        return;
                    }

                    // 2. แสดง HTML Table
                    if (response.html) {
                        document.getElementById('bulkPreviewContainer').innerHTML = response.html;
                    }

                    // 3. ตรวจสอบ Flag: canSubmit
                    // ถ้า True = ผ่านทุกรายการ -> แสดงปุ่ม Save
                    if (response.canSubmit === true) {
                        $('#divBulkActions').show();
                    } else {
                        $('#divBulkActions').hide();
                    }
                    
                    // ซ่อนผลลัพธ์เก่า (ถ้ามี)
                    $('#bulkResultContainer').hide();
                },
                error: function(err) {
                    showLoading(false);
                    alert('Error uploading file: ' + err.statusText);
                }
            });
        }

        function saveBulkData() {
            var dataPayload = [];
            
            // ดึงข้อมูลจาก Hidden Input (.row-data) ในตาราง Preview
            // (ใช้ .row-data แทน checkbox เพราะเราบังคับว่าต้องผ่านหมดถึงจะกดปุ่มได้)
            $('.row-data').each(function() {
                var val = $(this).val();
                if (val) {
                    dataPayload.push(JSON.parse(val));
                }
            });

            if (dataPayload.length === 0) {
                alert('No data found to process.');
                return;
            }

            if (!confirm('Are you sure you want to process ' + dataPayload.length + ' transactions? This will update SAP directly.')) {
                return;
            }

            showLoading(true, "Sending Batch to SAP... (Please wait)");

            $.ajax({
                url: 'Handler/SwitchUploadHandler.ashx?action=save',
                type: 'POST',
                data: { data: JSON.stringify(dataPayload) },
                success: function(res) {
                    showLoading(false);
                    if (res.success) {
                        alert(res.message);
                        renderBulkResults(res.results);
                        
                        // สำเร็จ -> ซ่อนปุ่ม Save และเคลียร์ Preview
                        $('#divBulkActions').hide();
                        $('#bulkPreviewContainer').html('<div class="alert alert-success">Transaction Completed Successfully!</div>');
                        document.getElementById('fileBulkSwitch').value = ""; // Clear file
                    } else {
                        // ไม่สำเร็จ (All-or-Nothing)
                        alert('Batch Failed: ' + res.message);
                        // อาจจะแสดง Error Detail ที่ด้านล่างก็ได้
                        $('#bulkResultContainer').show();
                        $('#tblBulkResult tbody').html('<tr><td colspan="7" class="text-danger text-center fw-bold">BATCH FAILED: ' + res.message + '</td></tr>');
                    }
                },
                error: function(err) {
                    showLoading(false);
                    alert('System Error: ' + err.statusText);
                }
            });
        }

        function renderBulkResults(results) {
            var tbody = $('#tblBulkResult tbody');
            tbody.empty();
            
            results.forEach(function(item, index) {
                var r = item.row; // Original data
                var cls = item.status === 'Success' ? '' : 'table-danger';
                var statusBadge = item.status === 'Success' ? '<span class="badge bg-success">Success</span>' : '<span class="badge bg-danger">Error</span>';
                
                var fromTxt = r.From.Year + '/' + r.From.Month + ' (' + r.From.Vendor + ')';
                var toTxt = r.Type === 'Switch' ? (r.To.Year + '/' + r.To.Month + ' (' + r.To.Vendor + ')') : '-';

                var html = `<tr class="${cls}">
                    <td>${index + 1}</td>
                    <td>${r.Type}</td>
                    <td>${fromTxt}</td>
                    <td>${toTxt}</td>
                    <td class="text-end">${parseFloat(r.Amount).toLocaleString('en-US', {minimumFractionDigits: 2})}</td>
                    <td>${statusBadge}</td>
                    <td>${item.message}</td>
                </tr>`;
                tbody.append(html);
            });

            $('#bulkResultContainer').show();
        }

        // ==========================================
        // Handle Switch Submit (Existing)
        // ==========================================
        async function handleSwitchSubmit(e) {
            e.preventDefault();
            clearValidationErrors();

            var formData = new FormData();
            formData.append('yearFrom', yearDropdownf.value); formData.append('monthFrom', monthDropdownf.value);
            formData.append('companyFrom', companyDropdownf.value); formData.append('categoryFrom', categoryDropdownf.value);
            formData.append('segmentFrom', segmentDropdownf.value); formData.append('brandFrom', brandDropdownf.value);
            formData.append('vendorFrom', vendorDropdownf.value);
            formData.append('yearTo', yearDropdownt.value); formData.append('monthTo', monthDropdownt.value);
            formData.append('companyTo', companyDropdownt.value); formData.append('categoryTo', categoryDropdownt.value);
            formData.append('segmentTo', segmentDropdownt.value); formData.append('brandTo', brandDropdownt.value);
            formData.append('vendorTo', vendorDropdownt.value);
            formData.append('amount', txtAmontSwitch.value);

            try {
                showLoading(true, "Validating...");
                var response = await fetch('Handler/ValidateHandler.ashx?action=validateSwitch', { method: 'POST', body: formData });
                var result = await response.json();
                showLoading(false);

                if (result.success) {
                    populatePreviewData();
                    var previewSwitchModal = new bootstrap.Modal(document.getElementById('previewSwitchModal'), { keyboard: false });
                    previewSwitchModal.show();
                } else {
                    showErrorModal(result.errors, 'Switch Transaction', '', result.availableBudget);
                }
            } catch (error) {
                showLoading(false);
                showSapErrorModal('Validation Error', 'Failed to validate data: ' + error.message);
            }
        }

        // ==========================================
        // Handle Extra Submit (Existing)
        // ==========================================
        async function handleExtraSubmit(e) {
            e.preventDefault();
            clearValidationErrors();

            var formData = new FormData();
            formData.append('year', yearDropdownE.value); formData.append('month', monthDropdownE.value);
            formData.append('company', companyDropdownE.value); formData.append('category', categoryDropdownE.value);
            formData.append('segment', segmentDropdownE.value); formData.append('brand', brandDropdownE.value);
            formData.append('vendor', vendorDropdownE.value);
            formData.append('amount', txtAmontEx.value);

            try {
                showLoading(true, "Validating...");
                var response = await fetch('Handler/ValidateHandler.ashx?action=validateExtra', { method: 'POST', body: formData });
                var result = await response.json();
                showLoading(false);

                if (result.success) {
                    populateExtraPreviewData();
                    var previewExtraModal = new bootstrap.Modal(document.getElementById('previewExtraModal'), { keyboard: false });
                    previewExtraModal.show();
                } else {
                    showErrorModal(result.errors, 'Extra Budget', 'Ex', result.currentBudget);
                }
            } catch (error) {
                showLoading(false);
                showSapErrorModal('Validation Error', 'Failed to validate data: ' + error.message);
            }
        }

        // ==========================================
        // [NEW FUNCTION] Show SAP Error Modal
        // ==========================================
        function showSapErrorModal(title, message) {
            document.getElementById('errorValidationModalTitle').textContent = title;
            document.getElementById('errorSummaryTitle').textContent = "An error occurred:";
            document.getElementById('errorSummaryText').textContent = "Please review the details below.";
            var errorListContainer = document.getElementById('errorListContainer');
            errorListContainer.innerHTML = `
                <div class="error-item">
                    <div class="error-item-icon"><i class="bi bi-exclamation-circle"></i></div>
                    <div class="error-item-content">
                        <div class="error-field-name"><span class="error-section-badge" style="background-color: #dc3545;">System</span>Error Details</div>
                        <p class="error-message">${message}</p>
                    </div>
                </div>`;
            if (!errorModal) errorModal = new bootstrap.Modal(document.getElementById('errorValidationModal'));
            errorModal.show();
        }

        // ==========================================
        // Show Validation Error Modal (Existing)
        // ==========================================
        function showErrorModal(errors, transactionType, suffix, availableBudget) {
            transactionType = transactionType || 'Transaction';
            suffix = suffix || '';
            document.getElementById('errorValidationModalTitle').textContent = 'Validation Error';
            document.getElementById('errorSummaryTitle').textContent = 'Please correct the following errors:';

            var fieldInfo = {
                'yearFrom': { name: 'Year', section: 'From', element: 'DDYearFrom' },
                'monthFrom': { name: 'Month', section: 'From', element: 'DDMonthFrom' },
                'companyFrom': { name: 'Company', section: 'From', element: 'DDCompanyFrom' },
                'categoryFrom': { name: 'Category', section: 'From', element: 'DDCategoryFrom' },
                'segmentFrom': { name: 'Segment', section: 'From', element: 'DDSegmentFrom' },
                'brandFrom': { name: 'Brand', section: 'From', element: 'DDBrandFrom' },
                'vendorFrom': { name: 'Vendor', section: 'From', element: 'DDVendorFrom' },
                'yearTo': { name: 'Year', section: 'To', element: 'DDYearTo' },
                'monthTo': { name: 'Month', section: 'To', element: 'DDMonthTo' },
                'companyTo': { name: 'Company', section: 'To', element: 'DDCompanyTo' },
                'categoryTo': { name: 'Category', section: 'To', element: 'DDCategoryTo' },
                'segmentTo': { name: 'Segment', section: 'To', element: 'DDSegmentTo' },
                'brandTo': { name: 'Brand', section: 'To', element: 'DDBrandTo' },
                'vendorTo': { name: 'Vendor', section: 'To', element: 'DDVendorTo' },
                'amount': { name: 'Amount (THB)', section: 'General', element: (transactionType === 'Extra Budget' ? 'txtAmontEx' : 'txtAmontSwitch') },
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
                    var budgetFormatted = parseFloat(availableBudget).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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
                    errorHtml += '<div class="error-item"><div class="error-item-icon"><i class="bi bi-exclamation-circle"></i></div><div class="error-item-content"><div class="error-field-name"><span class="error-section-badge" style="background-color: #6c757d;">General</span>Validation Error</div><p class="error-message">' + message + '</p></div></div>';
                } else if (info) {
                    var sectionColor = info.section === 'From' ? '#FF6B35' : info.section === 'To' ? '#4ECDC4' : info.section === 'Extra' ? '#28a745' : '#dc3545';
                    errorHtml += '<div class="error-item" data-field="' + info.element + '"><div class="error-item-icon"><i class="bi bi-x-circle"></i></div><div class="error-item-content"><div class="error-field-name"><span class="error-section-badge" style="background-color: ' + sectionColor + ';">' + info.section + '</span>' + info.name + '</div><p class="error-message">' + message + '</p></div></div>';
                    var element = document.getElementById(info.element);
                    if (element) element.classList.add('has-error');
                }
            }
            if (errorListContainer) errorListContainer.innerHTML = errorHtml;
            if (!errorModal) errorModal = new bootstrap.Modal(document.getElementById('errorValidationModal'));
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
                                setTimeout(function () { fieldElement.classList.remove('pulse-error'); }, 1000);
                            }, 300);
                        }
                    });
                })(errorItems[j]);
            }
        }

        // ==========================================
        // Helper Functions (Existing)
        // ==========================================
        function clearValidationErrors() {
            document.querySelectorAll('.has-error').forEach(el => el.classList.remove('has-error'));
            document.querySelectorAll('.pulse-error').forEach(el => el.classList.remove('pulse-error'));
        }

        function populatePreviewData() {
            const getSelectedText = (el) => el.options[el.selectedIndex]?.text || el.value;
            document.getElementById("tsYearFrom").value = getSelectedText(yearDropdownf);
            document.getElementById("tsMonthFrom").value = getSelectedText(monthDropdownf);
            document.getElementById("tsCompanyFrom").value = getSelectedText(companyDropdownf);
            document.getElementById("tsCategoryFrom").value = getSelectedText(categoryDropdownf);
            document.getElementById("tsSegmentFrom").value = getSelectedText(segmentDropdownf);
            document.getElementById("tsBrandFrom").value = getSelectedText(brandDropdownf);
            document.getElementById("tsVendorFrom").value = getSelectedText(vendorDropdownf);
            document.getElementById("tsYearTo").value = getSelectedText(yearDropdownt);
            document.getElementById("tsMonthTo").value = getSelectedText(monthDropdownt);
            document.getElementById("tsCompanyTo").value = getSelectedText(companyDropdownt);
            document.getElementById("tsCategoryTo").value = getSelectedText(categoryDropdownt);
            document.getElementById("tsSegmentTo").value = getSelectedText(segmentDropdownt);
            document.getElementById("tsBrandTo").value = getSelectedText(brandDropdownt);
            document.getElementById("tsVendorTo").value = getSelectedText(vendorDropdownt);
            document.getElementById("tsAmontSwitch").value = txtAmontSwitch.value;
        }

        function populateExtraPreviewData() {
            const getSelectedText = (el) => el.options[el.selectedIndex]?.text || el.value;
            document.getElementById("tsYearEx").value = getSelectedText(yearDropdownE);
            document.getElementById("tsMonthEx").value = getSelectedText(monthDropdownE);
            document.getElementById("tsCompanyEx").value = getSelectedText(companyDropdownE);
            document.getElementById("tsCategoryEx").value = getSelectedText(categoryDropdownE);
            document.getElementById("tsSegmentEx").value = getSelectedText(segmentDropdownE);
            document.getElementById("tsBrandEx").value = getSelectedText(brandDropdownE);
            document.getElementById("tsVendorEx").value = getSelectedText(vendorDropdownE);
            document.getElementById("tsAmontEx").value = txtAmontEx.value;
        }

        function showLoading(show, text = 'Processing...') {
            var overlay = document.getElementById('loadingOverlay');
            var loadingText = document.getElementById('loadingText');
            if (overlay) {
                if (show) {
                    if (loadingText) loadingText.textContent = text;
                    overlay.classList.add('active');
                } else {
                    overlay.classList.remove('active');
                }
            }
        }

        // ==========================================
        // Save Functions (Existing)
        // ==========================================
        async function saveSwitchingData() {
            showLoading(true, "Saving...");
            var currentUser = '<%= HttpUtility.JavaScriptStringEncode(Session("user").ToString()) %>';
            var formData = new FormData();
            formData.append('yearFrom', yearDropdownf.value); formData.append('monthFrom', monthDropdownf.value);
            formData.append('companyFrom', companyDropdownf.value); formData.append('categoryFrom', categoryDropdownf.value);
            formData.append('segmentFrom', segmentDropdownf.value); formData.append('brandFrom', brandDropdownf.value);
            formData.append('vendorFrom', vendorDropdownf.value);
            formData.append('yearTo', yearDropdownt.value); formData.append('monthTo', monthDropdownt.value);
            formData.append('companyTo', companyDropdownt.value); formData.append('categoryTo', categoryDropdownt.value);
            formData.append('segmentTo', segmentDropdownt.value); formData.append('brandTo', brandDropdownt.value);
            formData.append('vendorTo', vendorDropdownt.value);
            formData.append('amount', document.getElementById('tsAmontSwitch').value);
            formData.append('createdBy', currentUser || 'unknown');
            formData.append('remark', '');

            try {
                var response = await fetch('Handler/SaveOTBHandler.ashx?action=saveSwitching', { method: 'POST', body: formData });
                var result = await response.json();
                showLoading(false);
                if (result.success) {
                    bootstrap.Modal.getInstance(document.getElementById('previewSwitchModal')).hide();
                    alert(result.message || 'Save successful!');
                    // Reset Logic... (same as existing)
                    yearDropdownf.value = ""; monthDropdownf.value = ""; companyDropdownf.value = "";
                    txtAmontSwitch.value = "0.00";
                    InitMSData();
                } else {
                    bootstrap.Modal.getInstance(document.getElementById('previewSwitchModal')).hide();
                    showSapErrorModal('Save Failed (SAP Error)', result.message);
                }
            } catch (error) {
                showLoading(false);
                bootstrap.Modal.getInstance(document.getElementById('previewSwitchModal')).hide();
                showSapErrorModal('System Error', error.message);
            }
        }

        async function saveExtraData() {
            showLoading(true, "Saving...");
            var formData = new FormData();
            formData.append('year', yearDropdownE.value); formData.append('month', monthDropdownE.value);
            formData.append('company', companyDropdownE.value); formData.append('category', categoryDropdownE.value);
            formData.append('segment', segmentDropdownE.value); formData.append('brand', brandDropdownE.value);
            formData.append('vendor', vendorDropdownE.value);
            formData.append('amount', document.getElementById('tsAmontEx').value);
            formData.append('createdBy', 'System');
            formData.append('remark', '');

            try {
                var response = await fetch('Handler/SaveOTBHandler.ashx?action=saveExtra', { method: 'POST', body: formData });
                var result = await response.json();
                showLoading(false);
                if (result.success) {
                    bootstrap.Modal.getInstance(document.getElementById('previewExtraModal')).hide();
                    alert(result.message || 'Extra budget save successful!');
                    // Reset Logic...
                    yearDropdownE.value = ""; monthDropdownE.value = "";
                    txtAmontEx.value = "0.00";
                    InitMSData();
                } else {
                    bootstrap.Modal.getInstance(document.getElementById('previewExtraModal')).hide();
                    showSapErrorModal('Save Failed (SAP Error)', result.message);
                }
            } catch (error) {
                showLoading(false);
                bootstrap.Modal.getInstance(document.getElementById('previewExtraModal')).hide();
                showSapErrorModal('System Error', error.message);
            }
        }

        // ==========================================
        // Initialize Master Data
        // ==========================================
        var InitMSData = function () {
            InitSegment(segmentDropdownf); InitCategoty(categoryDropdownf); InitBrand(brandDropdownf); InitVendor(vendorDropdownf);
            InitMSYear(yearDropdownf); InitMonth(monthDropdownf); InitCompany(companyDropdownf);

            InitSegment(segmentDropdownt); InitCategoty(categoryDropdownt); InitBrand(brandDropdownt); InitVendor(vendorDropdownt);
            InitMSYear(yearDropdownt); InitMonth(monthDropdownt); InitCompany(companyDropdownt);

            InitSegment(segmentDropdownE); InitCategoty(categoryDropdownE); InitBrand(brandDropdownE); InitVendor(vendorDropdownE);
            InitMSYear(yearDropdownE); InitMonth(monthDropdownE); InitCompany(companyDropdownE);
        };

        // Master Data AJAX Functions (Simplified for brevity - use your existing ones)
        var InitSegment = function (el) { loadMasterData(el, 'SegmentMSList'); };
        var InitMSYear = function (el) { loadMasterData(el, 'YearMSList'); };
        var InitMonth = function (el) { loadMasterData(el, 'MonthMSList'); };
        var InitCompany = function (el) { loadMasterData(el, 'CompanyMSList'); };
        var InitCategoty = function (el) { loadMasterData(el, 'CategoryMSList'); };
        var InitBrand = function (el) { loadMasterData(el, 'BrandMSList'); };
        var InitVendor = function (el) { loadMasterData(el, 'VendorMSList'); };

        function loadMasterData(el, action) {
            if (!el) return;
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=' + action,
                type: 'POST',
                success: function (response) { el.innerHTML = response; }
            });
        }

        // Change Vendor Functions
        var changeVendorF = function () { updateVendor(segmentDropdownf, vendorDropdownf); };
        var changeVendorT = function () { updateVendor(segmentDropdownt, vendorDropdownt); };
        var changeVendorE = function () { updateVendor(segmentDropdownE, vendorDropdownE); };

        function updateVendor(segEl, vendEl) {
            if (!segEl || !vendEl) return;
            var segmentCode = segEl.value;
            if (!segmentCode) { InitVendor(vendEl); return; }
            var formData = new FormData(); formData.append('segmentCode', segmentCode);
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=VendorMSListChg', type: 'POST', data: formData, processData: false, contentType: false,
                success: function (response) { vendEl.innerHTML = response; }
            });
        }

        // Currency Formatting Functions (Existing)
        function formatCurrencyOnBlur(event) {
            const input = event.target;
            let value = parseFloat(input.value.replace(/,/g, ''));
            if (!isNaN(value)) {
                input.value = value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            } else { input.value = '0.00'; }
        }
        function cleanCurrencyOnFocus(event) {
            const input = event.target;
            let value = input.value.replace(/,/g, '');
            if (parseFloat(value) === 0) { input.value = ''; } else { input.value = value; }
            setTimeout(() => input.select(), 0);
        }
        function restrictToNumeric(event) {
            const key = event.key;
            if (['Backspace', 'Tab', 'Enter', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End', 'Delete'].includes(key)) return;
            if ((key === 'a' || key === 'c' || key === 'v' || key === 'x') && event.ctrlKey) return;
            if (key === '.' && !event.target.value.includes('.')) return;
            if (!/\d/.test(key)) event.preventDefault();
        }

        // Document Ready
        document.addEventListener('DOMContentLoaded', initial);
    </script>
</body>
</html>