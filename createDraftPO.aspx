<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="createDraftPO.aspx.vb" Inherits="BMS.createDraftPO" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - Create Draft PO</title>
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
            <li class="menu-item"><a href="default.aspx" class="menu-link"><i class="bi bi-box-arrow-left"></i> Logout</a></li>
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
                <h1 class="page-title" id="pageTitle">KBMS - Create Draft PO</h1>
            </div>
            <div class="user-info">
                <span class="d-none d-md-inline">Welcome, <%= HttpUtility.JavaScriptStringEncode(Session("fullname").ToString()) %></span>
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
                                    <input id="txtPONO" type="text" class="form-control" placeholder="Enter PO number" autocomplete="off">
                                </div>
                                <div class="form-group">
                                    <label>Amount (CCY)</label>
                                    <input id="txtAmtCCY" type="text" class="form-control" pattern="^\d+(\.\d{1,2})?$" title="Enter a valid amount (e.g., 123 or 123.45)" placeholder="0.00" autocomplete="off">
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
                                    <input id="txtExRate" type="text" class="form-control" pattern="^\d+(\.\d{1,2})?$" title="Enter a valid amount (e.g., 123 or 123.4523)" placeholder="0.0000" autocomplete="off">
                                </div>
                                <div class="form-group">
                                    <label>Remark</label>
                                    <input id="txtRemark" type="text" class="form-control" placeholder="Enter remark" autocomplete="off">
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
                                        <input id="tsPONO" type="text" class="form-control" placeholder="Enter PO number" autocomplete="off">
                                    </div>
                                    <div class="form-group">
                                        <label>Amount (CCY)</label>
                                        <input id="tsAmtCCY" type="text" class="form-control" placeholder="0.00" autocomplete="off">
                                    </div>
                                </div>

                                <!-- Row 6 -->
                                <div class="form-row-display form-row-item">
                                    <div class="form-group">
                                        <label>CCY</label>
                                        <input id="tsCCY" type="text" class="form-control" placeholder="Enter PO number" autocomplete="off">
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
                                        <input id="tsExRate" type="text" class="form-control" placeholder="0.00" pattern="^\d+(\.\d{1,2})?$" title="Enter a valid amount (e.g., 123 or 123.45)" autocomplete="off">
                                    </div>
                                    <div class="form-group">
                                        <label>Remark</label>
                                        <input id="tsRemark" type="text" class="form-control" placeholder="Enter remark" autocomplete="off">
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

    <div class="modal fade" id="uploadResultModal" tabindex="-1" aria-labelledby="uploadResultModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-xl">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="uploadResultModalLabel">
                    <i class="bi bi-clipboard-data"></i> Upload Results
                </h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <div id="uploadResultSummary" class="alert alert-info mb-3"></div>
                
                <div class="table-responsive" style="max-height: 500px; overflow-y: auto;">
                    <table class="table table-bordered table-hover" id="tblUploadResults">
                        <thead class="table-light sticky-top" style="top: 0; z-index: 1;">
                            <tr>
                                <th style="width: 50px;">#</th>
                                <th style="width: 150px;">Draft PO No.</th>
                                <th style="width: 100px;">Status</th>
                                <th>Message</th>
                            </tr>
                        </thead>
                        <tbody id="uploadResultBody">
                            </tbody>
                    </table>
                </div>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-primary" data-bs-dismiss="modal" onclick="closeAndRefresh()">OK & Refresh</button>
            </div>
        </div>
    </div>
</div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script>

        // ADDED: Modal instances
        let previewTxnModal, errorModal, successModal, uploadPreviewModal;
        let previewPOTXNModal;

        let uploadResultModal;

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
        let btnSubmit = document.getElementById('btnSubmit');

        let btnUpload = document.getElementById('btnUpload');
        let fileUploadInput = document.getElementById('fileUpload');

        let previewTableContainer = document.getElementById('previewTableContainer');
        let btnSubmitData = document.getElementById('btnSubmitData');


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
            let amtCCY = parseFloat(txtAmtCCY.value.replace(",", "")) || 0;
            let exRate = parseFloat(txtExRate.value.replace(",", "")) || 0;

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

            txtAmtTHB.value = (amtCCY * exRate).toLocaleString('en-US', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });
        }

        let initial = function () {
            const firstMenuLink = document.querySelector('.menu-link');
            if (firstMenuLink) {
                firstMenuLink.classList.add('expanded');
            }

            // *** ADD NEW EVENT LISTENERS FOR CURRENCY MASKING ***
            if (txtAmtTHB) {
                txtAmtTHB.addEventListener('keydown', restrictToNumeric);
                txtAmtTHB.addEventListener('focus', cleanCurrencyOnFocus);
                txtAmtTHB.addEventListener('blur', formatCurrencyOnBlur);

            }
            if (txtAmtCCY) {
                txtAmtCCY.addEventListener('keydown', restrictToNumeric);
                txtAmtCCY.addEventListener('focus', cleanCurrencyOnFocus);
                txtAmtCCY.addEventListener('blur', formatCurrencyOnBlur);
            }
            if (txtExRate) {
                txtExRate.addEventListener('keydown', restrictToNumeric);
            }

            if (ccyDropdown) {
                $(ccyDropdown).select2({
                    theme: "bootstrap-5"
                });
            }

            if (yearDropdown) {
                $(yearDropdown).select2({
                    theme: "bootstrap-5"
                });
            }

            if (monthDropdown) {
                $(monthDropdown).select2({
                    theme: "bootstrap-5"
                });
            }

            if (companyDropdown) {
                $(companyDropdown).select2({
                    theme: "bootstrap-5"
                });
            }

            if (segmentDropdown) {
                $(segmentDropdown).select2({
                    theme: "bootstrap-5"
                });
            }

            if (categoryDropdown) {
                $(categoryDropdown).select2({
                    theme: "bootstrap-5"
                });
            }

            if (brandDropdown) {
                $(brandDropdown).select2({
                    theme: "bootstrap-5"
                });
            }

            if (vendorDropdown) {
                $(vendorDropdown).select2({
                    theme: "bootstrap-5"
                });
            }

            //InitData master
            InitMSData();

            previewPOTXNModal = new bootstrap.Modal(document.getElementById('previewPOTXNModal'));
            errorModal = new bootstrap.Modal(document.getElementById('errorValidationModal'));
            successModal = new bootstrap.Modal(document.getElementById('successModal'));
            uploadPreviewModal = new bootstrap.Modal(document.getElementById('previewModal'));


            $('#DDSegment').on('select2:select',changeVendor);
            $('#DDVendor').on('select2:select',chengeCCY);
            $('#DDCCY').on('select2:select', currencyCal);
            txtAmtCCY.addEventListener('change', currencyCal);
            txtExRate.addEventListener('change', currencyCal);
            
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
            document.getElementById("tsAmtCCY").value = txtAmtCCY.value;
            document.getElementById("tsCCY").value = getSelectedText(ccyDropdown);
            document.getElementById("tsExRate").value = txtExRate.value;
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
            formData.append('pono', txtPONO.value);
            formData.append('amtCCY', txtAmtCCY.value);
            formData.append('ccy', ccyDropdown.value); // Send value
            formData.append('exRate', txtExRate.value);
            formData.append('amtTHB', txtAmtTHB.value);
            formData.append('remark', txtRemark.value);
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

            document.getElementById('txtPONO').value = "";
            document.getElementById('txtAmtCCY').value = "";
            document.getElementById('txtExRate').value = "";
            document.getElementById('txtAmtTHB').value = "0.00";
            document.getElementById('txtRemark').value = "";

            $('#DDYear').val(null).trigger('change');
            $('#DDMonth').val(null).trigger('change');
            $('#DDCompany').val(null).trigger('change');
            $('#DDCategory').val(null).trigger('change');
            $('#DDSegment').val(null).trigger('change');
            $('#DDBrand').val(null).trigger('change');
            $('#DDVendor').val(null).trigger('change');
            $('#DDCCY').val(null).trigger('change');

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

        // *** NEW: Upload Submit Handler ***
        function handleUploadSubmit(e) {
            e.preventDefault();
            var selectedRowsData = [];

            $('#previewTableContainer input[name="selectedRows"]:checked').each(function () {
                var cb = $(this);
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
                alert('Please select at least one valid row to save.');
                return;
            }

            var uploadBy = '<%= If(Session("user") IsNot Nothing, HttpUtility.JavaScriptStringEncode(Session("user").ToString()), "unknown") %>';
            var formData = new FormData();
            formData.append('uploadBy', uploadBy);
            formData.append('selectedData', JSON.stringify(selectedRowsData));

            showLoading(true, "Saving...", "Submitting selected rows...");
            btnSubmitData.disabled = true;

            $.ajax({
                url: 'Handler/POUploadHandler.ashx?action=savePreview',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                dataType: 'json',
                success: function (response) {
                    showLoading(false);
                    btnSubmitData.disabled = false;

                    uploadPreviewModal.hide();

                    previewTableContainer.innerHTML = '';
                    fileUploadInput.value = '';

                    showUploadResults(response);

                    // ตรวจสอบว่ามี error message กลับมาหรือไม่
                    //if (response.toLowerCase().includes("error") || response.includes("alert-danger")) {
                    //    showErrorSaveModal(response, 'Save Error');
                    //} else {
                    //    showSuccessModal('Success', response);
                    //}
                },
                error: function (xhr, status, error) {
                    showLoading(false);
                    btnSubmitData.disabled = false;
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

        // (NEW FUNCTION 1) Formats the number when user clicks away
        function formatCurrencyOnBlur(event) {
            const input = event.target;
            let value = parseFloat(input.value.replace(/,/g, '')); // Remove existing commas
            if (!isNaN(value)) {
                // Format to x,xxx.xx
                input.value = value.toLocaleString('en-US', {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2
                });
            } else {
                // Default to 0.00 if input is invalid
                input.value = '0.00';
            }
        }

        // (NEW FUNCTION 2) Clears formatting when user clicks in
        function cleanCurrencyOnFocus(event) {
            const input = event.target;
            let value = input.value.replace(/,/g, ''); // Remove commas

            // If the value is '0.00', clear it so user can type easily
            if (parseFloat(value) === 0) {
                input.value = '';
            } else {
                // Otherwise, just show the raw number
                input.value = value;
            }
            // Select the text for easy replacement
            setTimeout(() => input.select(), 0);
        }

        // (NEW FUNCTION 3) Prevents invalid characters (bound to 'keydown')
        function restrictToNumeric(event) {
            const input = event.target;
            const key = event.key;

            // Allow control keys (Backspace, Tab, Enter, Arrows, Home, End, Delete)
            if (['Backspace', 'Tab', 'Enter', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End', 'Delete'].includes(key)) {
                return;
            }

            // [BMS Gem Fix] Allow Ctrl+A, Ctrl+C, Ctrl+V, Ctrl+X
            if ((key === 'a' || key === 'c' || key === 'v' || key === 'x') && event.ctrlKey) {
                return;
            }

            // Allow one decimal point
            if (key === '.' && !input.value.includes('.')) {
                return;
            }

            // Allow only digits
            if (!/\d/.test(key)) {
                event.preventDefault(); // Block the key press
            }
        }

        // ============================================================
        // 2. เพิ่มฟังก์ชันใหม่สำหรับแสดงผลลัพธ์ราย Row (Show Result)
        // ============================================================
        function showUploadResults(results) {
            // 1. สรุปยอดรวม
            let total = results.length;
            let success = results.filter(r => r.Status === 'Success').length;
            let error = results.filter(r => r.Status === 'Error').length;

            let summaryHtml = `
            <strong>Process Complete:</strong> 
            Total <span class="badge bg-primary">${total}</span> | 
            Success <span class="badge bg-success">${success}</span> | 
            Failed <span class="badge bg-danger">${error}</span>
        `;

            // ปรับสี Alert ตามผลลัพธ์
            let alertClass = error > 0 ? 'alert-warning' : 'alert-success';
            $('#uploadResultSummary').removeClass('alert-info alert-success alert-warning alert-danger').addClass(alertClass).html(summaryHtml);

            // 2. สร้างตารางรายละเอียด
            let tbody = $('#uploadResultBody');
            tbody.empty();

            results.forEach((row, index) => {
                let isSuccess = row.Status === 'Success';
                let rowClass = isSuccess ? 'table-success' : 'table-danger';
                let icon = isSuccess ? '<i class="bi bi-check-circle-fill text-success"></i>' : '<i class="bi bi-x-circle-fill text-danger"></i>';

                let tr = `
                <tr class="${rowClass}">
                    <td class="text-center">${index + 1}</td>
                    <td>${row.DraftPONo || '-'}</td>
                    <td class="text-center">${icon} ${row.Status}</td>
                    <td>${row.Message}</td>
                </tr>
            `;
                tbody.append(tr);
            });

            // 3. เปิด Modal
            // (ต้องแน่ใจว่าประกาศตัวแปร uploadResultModal ไว้แล้วใน initial)
            if (!uploadResultModal) {
                uploadResultModal = new bootstrap.Modal(document.getElementById('uploadResultModal'));
            }
            uploadResultModal.show();
        }

        // ฟังก์ชันเสริม: กดปุ่ม OK ใน Result Modal แล้วให้ Refresh หน้าจอหรือเคลียร์ค่า
        function closeAndRefresh() {
            // เคลียร์ค่า Preview
            document.getElementById('previewTableContainer').innerHTML = '';
            document.getElementById('fileUpload').value = '';

            // (Optional) โหลดข้อมูลในตารางหลักใหม่ เพื่อให้เห็นรายการที่เพิ่ง Save
            // search(); 
            // หรือถ้าต้องการแค่ปิด Modal เฉยๆ ก็ไม่ต้องทำอะไร
        }


        // Initialize
        document.addEventListener('DOMContentLoaded', initial);
    </script>
</body>
</html>
