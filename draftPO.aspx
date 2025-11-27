<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="draftPO.aspx.vb" Inherits="BMS.draftPO" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - Draft PO</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="style/theme.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
</head>
<body>
    <!-- Sidebar Overlay -->
    <div class="sidebar-overlay" id="sidebarOverlay" onclick="toggleSidebar()"></div>

    <!-- Sidebar -->
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
                <h1 class="page-title" id="pageTitle">KBMS - Draft PO</h1>
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
                                <div class="row g-3 mb-3">
                                    <div class="col-12 col-md-6">
                                        <label for="ddYearEdit">Year</label>
                                        <select id="ddYearEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="year"></div>
                                    </div>
                                    <div class="col-12 col-md-6">
                                        <label for="ddCategoryEdit">Category</label>
                                        <select id="ddCategoryEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="category"></div>
                                    </div>
                                </div>

                                <!-- Row 2 -->
                                <div class="row g-3 mb-3">
                                    <div class="col-12 col-md-6">
                                        <label for="ddMonthEdit">Month</label>
                                        <select id="ddMonthEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="month"></div>
                                    </div>
                                    <div class="col-12 col-md-6">
                                        <label for="ddSegmentEdit">Segment</label>
                                        <select id="ddSegmentEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="segment"></div>
                                    </div>
                                </div>

                                <!-- Row 3 -->
                                <div class="row g-3 mb-3">
                                    <div class="col-12 col-md-6">
                                        <label for="ddCompanyEdit">Company</label>
                                        <select id="ddCompanyEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="company"></div>
                                    </div>
                                    <div class="col-12 col-md-6">
                                        <label for="ddBrandEdit">Brand</label>
                                        <select id="ddBrandEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="brand"></div>
                                    </div>
                                </div>

                                <!-- Row 4 -->
                                <div class="row g-3 mb-3">
                                    <div class="col-12 col-md-6">
                                        <label for="txtPONOEdit">Draft PO no.</label>
                                        <input id="txtPONOEdit" type="text" class="form-control" autocomplete="off">
                                        <div class="validation-message" data-field="pono"></div>
                                    </div>
                                    <div class="col-12 col-md-6">
                                        <label for="ddVendorEdit">Vendor</label>
                                        <select id="ddVendorEdit" class="form-select">
                                        </select>
                                        <div class="validation-message" data-field="vendor"></div>
                                    </div>
                                </div>

                                <!-- Row 5 -->
                                <div class="row g-3 mb-3">
                                    <div class="col-12 col-md-6">
                                        <label for="txtAmtCCYEdit">Amount (CCY)</label>
                                        <input id="txtAmtCCYEdit" type="text" class="form-control" placeholder="0.00" autocomplete="off">
                                        <div class="validation-message" data-field="amtCCY"></div>
                                    </div>
                                    <div class="col-12 col-md-6">
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
                                <div class="row g-3 mb-3">
                                    <div class="col-12 col-md-6">
                                        <label for="txtExRateEdit">Exchange rate</label>
                                        <input id="txtExRateEdit" type="text" class="form-control" placeholder="0.00" autocomplete="off">
                                        <div class="validation-message" data-field="exRate"></div>
                                    </div>
                                    <div class="col-12 col-md-6">
                                        <label for="txtAmtTHBEdit">Amount (THB)</label>
                                        <input id="txtAmtTHBEdit" type="text" class="form-control" readonly style="background: #e9ecef;" autocomplete="off">
                                    </div>
                                </div>
                                
                                <!-- Row 7 -->
                                <div class="row g-3 mb-3">
                                    <div class="col-12 col-md-6" style="flex: 1;">
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
            ddCCYEdit = document.getElementById('ddCCYEdit');

            txtPONOEdit = document.getElementById('txtPONOEdit');
            txtAmtCCYEdit = document.getElementById('txtAmtCCYEdit');
            
            txtExRateEdit = document.getElementById('txtExRateEdit');
            txtAmtTHBEdit = document.getElementById('txtAmtTHBEdit');
            txtRemarkEdit = document.getElementById('txtRemarkEdit');
            btnExport = document.getElementById('btnExport');
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
            btnExport.addEventListener('click', exportTXN);
            // Add Listeners for Edit Modal Calculations
            txtAmtCCYEdit.addEventListener('input', currencyCalEdit);
            txtExRateEdit.addEventListener('input', currencyCalEdit);
           
            $('#ddCCYEdit').on('select2:select', currencyCalEdit);
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

            $('#ddYearFilter').val(null).trigger('change');
            $('#ddMonthFilter').val(null).trigger('change');
            $('#ddCompanyFilter').val(null).trigger('change');
            $('#ddCategoryFilter').val(null).trigger('change');
            $('#ddSegmentFilter').val(null).trigger('change');
            $('#ddBrandFilter').val(null).trigger('change');
            $('#ddVendorFilter').val(null).trigger('change');
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


                // [BMS Gem Fix]: ตรวจสอบสถานะเพื่อจัดการปุ่ม Edit
                const status = (row.Status || '').trim(); // ตัดช่องว่างเผื่อมี
                const statusLower = status.toLowerCase();
                const isMatched = statusLower === 'matched';
                const isCancelled = statusLower === 'cancelled';

                const canDelete = (statusLower === 'draft' || statusLower === 'matching' || statusLower === 'edited');

                // กำหนด Class สีสถานะ
                let statusClass = '';
                if (isMatched) statusClass = 'table-success';
                else if (isCancelled) statusClass = 'text-danger';

                // กำหนดสถานะปุ่ม Edit
                // ปิดปุ่มถ้าเป็น Matched หรือ Cancelled
                const isEditDisabled = isMatched || isCancelled;
                const editBtnClass = isEditDisabled ? 'btn-secondary' : 'btn-action'; // สีเทาถ้า disabled, สีเหลืองถ้าปกติ
                const editBtnAttr = isEditDisabled ? 'disabled' : '';
                const editBtnTitle = isEditDisabled ? 'Cannot edit because status is ' + status : 'Edit ' + (row.DraftPO_No || '');


                // Helper to format numbers
                const formatNum = (num, decimals = 2) => (num != null ? parseFloat(num).toFixed(decimals) : '0.00');

                return `
                    <tr class="${statusClass}" >
                        <td  class="${statusClass}" >${createdDate}</td>
                        <td  class="${statusClass}" >${row.DraftPO_No || ''}</td>
                        <td  class="${statusClass}" >${row.PO_Type || ''}</td>
                        <td  class="${statusClass}" >${row.PO_Year || ''}</td>
                        <td  class="${statusClass}" >${row.PO_Month_Name || ''}</td>
                        <td  class="${statusClass}" >${row.Category_Code || ''}</td>
                        <td  class="${statusClass}" >${row.Category_Name || ''}</td>
                        <td  class="${statusClass}" >${row.Company_Code || ''}</td>
                        <td  class="${statusClass}" >${row.Segment_Code || ''}</td>
                        <td  class="${statusClass}" >${row.Segment_Name || ''}</td>
                        <td  class="${statusClass}" >${row.Brand_Code || ''}</td>
                        <td  class="${statusClass}" >${row.Brand_Name || ''}</td>
                        <td  class="${statusClass}" >${row.Vendor_Code || ''}</td>
                        <td  class="${statusClass}" >${row.Vendor_Name || ''}</td>
                        <td class="${statusClass} text-end">${formatNum(row.Amount_THB)}</td>
                        <td class="${statusClass} text-end">${formatNum(row.Amount_CCY)}</td>
                        <td class="${statusClass}" >${row.CCY || ''}</td>
                        <td  class="${statusClass} text-end">${formatNum(row.Exchange_Rate, 4)}</td>
                        <td  class="${statusClass}" >${row.Actual_PO_Ref || ''}</td>
                        <td  class="${statusClass}" >${(row.Status === 'Matched' ? 'Matched' : (row.Status || '')) }</td>
                        <td  class="${statusClass}" >${statusDate}</td>
                        <td  class="${statusClass}" >${row.Remark || ''}</td>
                        <td  class="${statusClass}" >${row.Status_By || ''}</td>
                        <td  class="${statusClass}" >
                            <div class="d-flex gap-1">
                            <button class="btn ${editBtnClass} btn-edit-po" 
                                    data-draftpoid="${row.DraftPO_ID}" 
                                    title="${editBtnTitle}"
                                    ${editBtnAttr}>
                                <i class="bi bi-pencil"></i> Edit
                            </button>
                            ${canDelete ? `
                            <button class="btn btn-danger btn-sm btn-delete-po" 
                                    data-draftpoid="${row.DraftPO_ID}" 
                                    data-pono="${row.DraftPO_No}"
                                    title="Delete (Cancel)">
                                <i class="bi bi-trash"></i>
                            </button>
                            ` : ''}
                            </div>
                        </td>
                    </tr>
                `;
            }).join('');

            draftPOTableBody.innerHTML = rowsHtml;
            addEditButtonListeners();
            addDeleteButtonListeners();
        }

        // ==========================================
        // Edit PO Functions
        // ==========================================
        function addDeleteButtonListeners() {
            document.querySelectorAll('.btn-delete-po').forEach(button => {
                button.addEventListener('click', handleDeleteClick);
            });
        }
        function addEditButtonListeners() {
            document.querySelectorAll('.btn-edit-po').forEach(button => {
                button.addEventListener('click', handleEditClick);
            });
        }

        async function handleDeleteClick(event) {
            const btn = event.currentTarget;
            const draftPOID = btn.dataset.draftpoid;
            const poNo = btn.dataset.pono;

            if (!confirm(`Are you sure you want to CANCEL Draft PO No: ${poNo}?`)) {
                return;
            }

            showLoading(true, "Cancelling...");

            const formData = new FormData();
            formData.append('draftPOID', draftPOID);

            try {
                const response = await fetch('Handler/DataPOHandler.ashx?action=deleteDraftPO', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

                const result = await response.json();
                showLoading(false);

                if (result.success) {
                    alert(result.message);
                    // Refresh Table
                    handleViewData();
                } else {
                    alert('Error: ' + result.message);
                }

            } catch (error) {
                showLoading(false);
                console.error('Delete error:', error);
                alert('Fatal error cancelling Draft PO: ' + error.message);
            }
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

            $('#ddYearEdit').val(data.PO_Year).trigger('change');
            $('#ddMonthEdit').val(data.PO_Month).trigger('change');
            $('#ddCompanyEdit').val(data.Company_Code).trigger('change');
            $('#ddCategoryEdit').val(data.Category_Code).trigger('change');
            $('#ddSegmentEdit').val(data.Segment_Code).trigger('change');
            $('#ddBrandEdit').val(data.Brand_Code).trigger('change');
            $('#ddVendorEdit').val(data.Vendor_Code).trigger('change');
            $('#ddCCYEdit').val(data.CCY).trigger('change');

            txtPONOEdit.value = data.DraftPO_No;
            txtAmtCCYEdit.value = parseFloat(data.Amount_CCY || 0).toFixed(2);
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
            formData.append('pono', txtPONOEdit.value); 
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

        // ฟังก์ชันสำหรับ Export Excel
        let exportTXN = function () {
            console.log("Export TXN clicked");

            // รวบรวมค่าจาก Dropdown Filter ทั้งหมด
            var params = new URLSearchParams();
            params.append('action', 'exportdraftpo'); // Action ที่ Handler รอรับ

            // Map ID ของ Dropdown ให้ตรงกับ Parameter ที่ Handler ต้องการ
            params.append('year', ddYearFilter.value);
            params.append('month', ddMonthFilter.value);
            params.append('company', ddCompanyFilter.value);
            params.append('category', ddCategoryFilter.value);
            params.append('segment', ddSegmentFilter.value);
            params.append('brand', ddBrandFilter.value);
            params.append('vendor', ddVendorFilter.value);

           

            // ใช้ window.location เพื่อเรียก Download ไฟล์ (GET Request)
            window.location.href = 'Handler/DataPOHandler.ashx?' + params.toString();
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