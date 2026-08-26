<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="draftOTB.aspx.vb" Inherits="BMS.draftOTB" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-g">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - draft OTB</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="style/theme.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
</head>
<body>
    <form id="mainForm" action="/" method="post">

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
        <li class="menu-item" id="grpmenuOTBPlan" runat="server">
            <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'otbPlan')">
                <i class="bi bi-clipboard-data"></i>
                <span>OTB Plan / Revise</span>
                <i class="bi bi-chevron-down"></i>
            </a>
            <ul class="submenu" id="otbPlan">
                <li  id="menuDraftOTBPlan" runat="server" ><a href="draftOTB.aspx" class="menu-link active">Draft OTB Plan</a></li>
                <li id="menuApprovedOTBPlan" runat="server" ><a href="approvedOTB.aspx" class="menu-link">Approved OTB Plan</a></li>
            </ul>
        </li>
        <li class="menu-item" id="grpmenuOTBSwitching" runat="server">
            <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'otbSwitching')">
                <i class="bi bi-arrow-left-right"></i>
                <span>OTB Switching</span>
                <i class="bi bi-chevron-down"></i>
            </a>
            <ul class="submenu" id="otbSwitching">
               <li id="menuCreateOTBSwitching" runat="server" ><a href="createOTBswitching.aspx" class="menu-link">Create OTB Switching</a></li>
                <li id="menuSwitchingTransaction" runat="server" ><a href="transactionOTBSwitching.aspx" class="menu-link">Switching Transaction</a></li>
            </ul>
        </li>
        <li class="menu-item"  id="grpmenuPO" runat="server" >
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
        <li class="menu-item" id="menuOTBRemaining" runat="server">
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
        <li class="menu-item" id="grpmenuAdmin" runat="server">
            <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'adminTools')">
                <i class="bi bi-shield-lock"></i>
                <span>Admin</span>
                <i class="bi bi-chevron-down"></i>
            </a>
            <ul class="submenu" id="adminTools">
                <li id="menuAdminMatchPO" runat="server"><a href="admin_matchPO.aspx" class="menu-link">Admin Match PO</a></li>
                <li id="menuManageUsers" runat="server"><a href="manage_users.aspx" class="menu-link">Manage Users</a></li>
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
                <button class="menu-toggle" type="button" onclick="toggleSidebar()">
                    <i class="bi bi-list"></i>
                </button>
                <h1 class="page-title" id="pageTitle">KBMS - Draft OTB</h1>
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
            <div class="col-md-3 col-lg-2 mt-1 mt-md-0">
                <button id="btnUpload" class="btn btn-upload btn-custom w-150" type="button">
                    <i class="bi bi-upload"></i> Upload Draft OTB
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
                                <option value=''>-- กรุณาเลือก Type --</option>
                                <option value="Original" >Original</option> 
                                <option value="Revise" >Revise</option>
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
                <button type="button"  class="btn btn-export btn-custom" style="display:none;" id="btnExportSUM">
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
                                <th>Current Approved</th>
                                <th>TO-BE Amount (THB)</th>
                                <th>Diff</th>
                                <th>Status</th>
                                <th>Version</th>
                                <th>Remark</th>
                                <th>CreateBy</th>
                            </tr>
                        </thead>
                        <tbody id="tableViewBody">
                        </tbody>
                    </table>
                </div>
            </div>


            <div class="approval-buttons mt-4">
                <button type="button" class="btn btn-success" id="btnApprove">
                    <i class="bi bi-check-circle"></i> Approve Selected
                </button>
                 <button type="button" class="btn btn-danger" id="btnDelete">
                     <i class="bi bi-x-circle"></i> Delete Selected
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

    <div id="draftOtbActiveJobBar" class="alert alert-primary d-none mt-3" role="status" aria-live="polite">
        <div class="d-flex flex-wrap align-items-center justify-content-between gap-2">
            <span>A Draft OTB background job is available. You can reopen its progress at any time.</span>
            <div class="d-flex flex-wrap gap-2">
                <button type="button" id="btnOpenUploadJob" class="btn btn-sm btn-outline-primary d-none">Open upload progress</button>
                <button type="button" id="btnOpenApprovalJob" class="btn btn-sm btn-outline-primary d-none">Open approval progress</button>
            </div>
        </div>
    </div>

    <div class="loading-overlay" id="loadingOverlay">
    <div class="loading-content">
        <div class="loading-spinner"></div>
        <p class="loading-text">กำลังโหลดข้อมูล...</p>
        <p class="loading-subtext">กรุณารอสักครู่</p>
    </div>
</div>

    <div class="modal fade" id="draftOtbJobModal" tabindex="-1" aria-labelledby="draftOtbJobModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-centered">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="draftOtbJobModalLabel">Draft OTB progress</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div id="draftOtbJobIdentity" class="alert alert-light border py-2 d-none">
                        <div class="d-flex flex-wrap align-items-center justify-content-between gap-2">
                            <div><strong>Job ID:</strong> <code id="draftOtbJobId"></code></div>
                            <button type="button" id="btnCopyDraftOtbJobId" class="btn btn-sm btn-outline-secondary">Copy Job ID</button>
                        </div>
                    </div>
                    <div id="draftOtbJobSwitcher" class="btn-group btn-group-sm mb-3 d-none" role="group" aria-label="Select active Draft OTB job">
                        <button type="button" id="btnViewUploadJob" class="btn btn-outline-primary d-none">View upload</button>
                        <button type="button" id="btnViewApprovalJob" class="btn btn-outline-primary d-none">View approval</button>
                    </div>
                    <div class="d-flex justify-content-between mb-2">
                        <strong id="draftOtbJobStage">Queued</strong>
                        <strong id="draftOtbJobPercent">0%</strong>
                    </div>
                    <div class="progress" style="height:24px;" role="progressbar" aria-label="Draft OTB progress" aria-valuemin="0" aria-valuemax="100">
                        <div id="draftOtbJobProgressBar" class="progress-bar progress-bar-striped progress-bar-animated" style="width:0%;">0%</div>
                    </div>
                    <div id="draftOtbJobCounts" class="small text-muted mt-3">Processed 0 / 0 rows</div>
                    <div id="draftOtbJobMessage" class="alert alert-info mt-3 mb-0">The server is preparing this job.</div>
                    <div id="draftOtbCloseSafety" class="small text-muted mt-2">Keep this page open until the server returns a Job ID.</div>
                    <div id="draftOtbReconciliationDetails" class="d-none mt-3">
                        <div id="draftOtbReconciliationSummary" class="alert alert-warning py-2"></div>
                        <div class="d-flex flex-wrap gap-2 mb-2">
                            <button type="button" id="btnReviewReconciliationResults" class="btn btn-sm btn-outline-warning">I have reviewed these results</button>
                            <button type="button" id="btnDownloadReconciliationResults" class="btn btn-sm btn-outline-danger d-none">Download all results (JSON)</button>
                        </div>
                        <div id="draftOtbReconciliationResultTable" class="table-responsive" style="max-height:360px; overflow:auto;"></div>
                    </div>
                </div>
                <div class="modal-footer">
                    <a id="btnDownloadUploadErrors" class="btn btn-outline-danger d-none" href="#">
                        <i class="bi bi-download"></i> Download all errors
                    </a>
                    <button type="button" id="btnSaveValidatedUpload" class="btn btn-primary d-none">
                        <i class="bi bi-database-check"></i> Save all validated rows
                    </button>
                    <button type="button" id="btnAcknowledgeReconciliation" class="btn btn-danger d-none">I have recorded this Job ID</button>
                    <button type="button" id="btnAcknowledgeValidation" class="btn btn-warning d-none">I have reviewed the validation errors</button>
                    <button type="button" class="btn btn-secondary job-modal-close" data-bs-dismiss="modal">Close</button>
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="approvalPreviewModal" tabindex="-1" aria-labelledby="approvalPreviewModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="approvalPreviewModalLabel">Preview before sending to SAP</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div class="alert alert-warning">
                        Review the current approval batch grouped by Company / Year / Month / Category before sending it to SAP.
                    </div>
                    <div id="approvalPreviewTableContainer" class="table-responsive" style="max-height:600px; overflow:auto;"></div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="button" id="btnConfirmApprovalJob" class="btn btn-success">
                        <i class="bi bi-check-circle"></i> Confirm and start approval
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- =================================================== -->
    <!-- ===== NEW: APPROVAL RESULT MODAL ================== -->
    <!-- =================================================== -->
    <div class="modal fade" id="approvalResultModal" tabindex="-1" aria-labelledby="approvalResultModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="approvalResultModalLabel">Approval Results</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div id="approvalResultSummary" class="alert alert-info"></div>
                    <div id="approvalResultTableContainer" class="table-responsive" style="max-height:600px; overflow:auto;">
                        <!-- Table will be injected here by JS -->
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" id="btnDownloadApprovalResults" class="btn btn-outline-danger d-none">
                        <i class="bi bi-download"></i> Download all results (JSON)
                    </button>
                    <button type="button" class="btn btn-primary" data-bs-dismiss="modal">OK</button>
                </div>
            </div>
        </div>
    </div>
    <!-- =================================================== -->
    <!-- ===== END: NEW MODAL ============================== -->
    <!-- =================================================== -->

</body>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="script/bms-loading.js"></script>
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
    let btnExportSUM = document.getElementById("btnExportSUM");
    let btnApprove = document.getElementById('btnApprove');
    let btnDelete = document.getElementById('btnDelete');
    let btnSelectAll = document.getElementById('btnSelectAll');
    let btnDeselectAll = document.getElementById('btnDeselectAll');

    let approvalResultModal;
    let draftOtbJobModal;
    let approvalPreviewModal;
    let pendingApprovalRunNos = [];
    let pendingApprovalPreviewHash = '';
    let uploadPollTimer = null;
    let approvalPollTimer = null;
    let activeUploadJobId = null;
    let activeApprovalJobId = null;
    let activeUploadStatus = null;
    let activeApprovalStatus = null;
    let displayedJobType = null;
    let currentReconciliation = null;
    let currentApprovalResults = null;
    let startRequestInFlight = false;
    let jobDiscoveryComplete = false;
    const approvalPreviewDisplayLimit = 500;
    const draftOtbRequestTimeoutMs = 30000;
    const draftOtbStartTimeoutMs = 120000;
    const draftOtbCurrentUser = '<%= HttpUtility.JavaScriptStringEncode(Session("user").ToString()) %>' || 'unknown';
    const uploadStorageKey = 'bmsDraftOtbUploadJob:' + draftOtbCurrentUser;
    const approvalStorageKey = 'bmsDraftOtbApprovalJob:' + draftOtbCurrentUser;


    $(document).ready(function () {
        draftOtbJobModal = new bootstrap.Modal(document.getElementById('draftOtbJobModal'));
        approvalPreviewModal = new bootstrap.Modal(document.getElementById('approvalPreviewModal'));
        approvalResultModal = new bootstrap.Modal(document.getElementById('approvalResultModal'));

        $('#btnUpload').on('click', function (e) {
            e.preventDefault();
            if (!ensureNoBlockingJob('start another upload')) return;
            const uploadButton = this;
            const fileInput = $('#fileUpload')[0];
            const file = fileInput.files[0];

            if (!file) {
                alert('Please select a file.');
                return;
            }
            if (file.size > 20 * 1024 * 1024) {
                alert('The selected file exceeds the 20 MB limit.');
                return;
            }

            uploadButton.disabled = true;
            startRequestInFlight = true;
            resetJobModal('Uploading Draft OTB file', 'upload');
            draftOtbJobModal.show();
            renderJobProgress({
                stage: 'Preparing file signature', progress: 0, processedRows: 0, totalRows: 0,
                message: 'Calculating a stable signature so a retry cannot create a duplicate job.'
            }, 'upload');

            getUploadFileSignature(file).then(function (fileSignature) {
                startUploadRequest(file, fileSignature, uploadButton);
            }).catch(function (error) {
                uploadButton.disabled = false;
                startRequestInFlight = false;
                renderJobFailure('Unable to prepare the selected file: ' + (error.message || error));
            });
        });

        $('#btnSaveValidatedUpload').on('click', function () {
            if (!activeUploadJobId) return;
            if (!confirm('Save every validated row in this file? The save is all-or-nothing.')) return;
            const button = this;
            button.disabled = true;
            const formData = new FormData();
            formData.append('jobId', activeUploadJobId);
            formData.append('uploadBy', draftOtbCurrentUser);
            $.ajax({
                url: 'Handler/UploadHandler.ashx?action=saveUploadJob',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                dataType: 'json',
                timeout: draftOtbRequestTimeoutMs,
                success: function (response) {
                    button.disabled = false;
                    if (!response.success) {
                        renderJobFailure(response.message || 'Unable to start save job.');
                        return;
                    }
                    $('#btnSaveValidatedUpload').addClass('d-none');
                    renderJobProgress(response, 'upload');
                    pollUploadJob(activeUploadJobId);
                },
                error: function (xhr, status, error) {
                    button.disabled = false;
                    $('#btnSaveValidatedUpload').addClass('d-none');
                    showPollingWarning('upload');
                    pollUploadJob(activeUploadJobId);
                }
            });
        });

        $('#btnConfirmApprovalJob').on('click', startConfirmedApprovalJob);
        $('#btnViewUploadJob').on('click', function () { showStoredJob('upload'); });
        $('#btnViewApprovalJob').on('click', function () { showStoredJob('approval'); });
        $('#btnOpenUploadJob').on('click', function () { showStoredJob('upload'); });
        $('#btnOpenApprovalJob').on('click', function () { showStoredJob('approval'); });
        $('#btnCopyDraftOtbJobId').on('click', copyDisplayedJobId);
        $('#btnAcknowledgeReconciliation').on('click', acknowledgeReconciliation);
        $('#btnAcknowledgeValidation').on('click', acknowledgeValidationFailure);
        $('#btnReviewReconciliationResults').on('click', markReconciliationReviewed);
        $('#btnDownloadReconciliationResults').on('click', downloadReconciliationResults);
        $('#btnDownloadApprovalResults').on('click', downloadApprovalResults);
        $('#draftOtbJobModal').on('hide.bs.modal', function (event) {
            if (currentReconciliation) {
                event.preventDefault();
                $('#draftOtbJobMessage').removeClass('alert-info alert-success alert-warning').addClass('alert-danger')
                    .text('This job may already have changed SAP. Record the Job ID and acknowledge it before closing.');
            }
        });

        restoreJobsOnLoad();
    });

    function createClientRequestId(prefix) {
        if (window.crypto && window.crypto.randomUUID) {
            return prefix + '-' + window.crypto.randomUUID();
        }
        return prefix + '-' + Date.now() + '-' + Math.random().toString(16).slice(2);
    }

    function readJobState(key) {
        const raw = localStorage.getItem(key);
        if (!raw) return null;
        try {
            const parsed = JSON.parse(raw);
            return parsed && typeof parsed === 'object' ? parsed : null;
        } catch (error) {
            return { jobId: raw, status: null, updatedAt: null };
        }
    }

    function writeJobState(key, state) {
        localStorage.setItem(key, JSON.stringify(state));
    }

    function getOrCreateRequestState(key, prefix, signature) {
        const existing = readJobState(key);
        if (existing && !existing.jobId && existing.clientRequestId && existing.signature === signature) {
            return existing;
        }
        const state = {
            clientRequestId: createClientRequestId(prefix),
            signature: signature,
            jobId: null,
            status: 'Starting',
            updatedAt: new Date().toISOString()
        };
        // Persist before the request. A retry after a timeout will therefore use
        // the same idempotency key instead of creating a duplicate server job.
        writeJobState(key, state);
        return state;
    }

    function getUploadFileSignature(file) {
        const metadata = [file.name || '', file.size || 0, file.lastModified || 0, file.type || ''].join('|');
        return readUploadFileBuffer(file).then(function (buffer) {
            function fallbackContentHash() {
                // Web Crypto requires a secure browser context. Keep a
                // content-bound fallback for internal HTTP test environments.
                const bytes = new Uint8Array(buffer);
                let hashA = 2166136261;
                let hashB = 2246822519;
                for (let index = 0; index < bytes.length; index += 1) {
                    hashA = Math.imul(hashA ^ bytes[index], 16777619) >>> 0;
                    hashB = Math.imul(hashB ^ bytes[index], 3266489917) >>> 0;
                }
                return 'content32x2|' + hashA.toString(16).padStart(8, '0') + hashB.toString(16).padStart(8, '0') + '|' + metadata;
            }
            if (window.crypto && window.crypto.subtle) {
                return window.crypto.subtle.digest('SHA-256', buffer).then(function (digest) {
                    const hash = Array.prototype.map.call(new Uint8Array(digest), function (value) {
                        return value.toString(16).padStart(2, '0');
                    }).join('');
                    return 'sha256|' + hash + '|' + metadata;
                }).catch(fallbackContentHash);
            }
            return fallbackContentHash();
        });
    }

    function readUploadFileBuffer(file) {
        if (file.arrayBuffer) return file.arrayBuffer();
        return new Promise(function (resolve, reject) {
            const reader = new FileReader();
            reader.onload = function () { resolve(reader.result); };
            reader.onerror = function () { reject(reader.error || new Error('Unable to read the selected file.')); };
            reader.readAsArrayBuffer(file);
        });
    }

    function startUploadRequest(file, fileSignature, uploadButton) {
        const requestState = getOrCreateRequestState(uploadStorageKey, 'upload', fileSignature);
        const formData = new FormData();
        formData.append('file', file);
        formData.append('uploadBy', draftOtbCurrentUser);
        formData.append('clientRequestId', requestState.clientRequestId);

        $.ajax({
            url: 'Handler/UploadHandler.ashx?action=startUploadJob',
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            dataType: 'json',
            timeout: draftOtbStartTimeoutMs,
            xhr: function () {
                const xhr = $.ajaxSettings.xhr();
                if (xhr.upload) {
                    xhr.upload.addEventListener('progress', function (event) {
                        if (event.lengthComputable) {
                            const acceptedPercent = Math.min(5, Math.floor((event.loaded / event.total) * 5));
                            renderJobProgress({
                                stage: 'Uploading file to server', progress: acceptedPercent,
                                processedRows: 0, totalRows: 0,
                                message: 'The server will validate every row after the upload is accepted.'
                            }, 'upload');
                        }
                    });
                }
                return xhr;
            },
            success: function (response) {
                uploadButton.disabled = false;
                startRequestInFlight = false;
                if (!response.success) {
                    renderJobFailure(response.message || 'Unable to start upload job.');
                    return;
                }
                activeUploadJobId = response.jobId;
                activeUploadStatus = response.status || 'Queued';
                displayedJobType = 'upload';
                requestState.jobId = activeUploadJobId;
                requestState.status = activeUploadStatus;
                requestState.updatedAt = response.updatedAt || new Date().toISOString();
                writeJobState(uploadStorageKey, requestState);
                renderJobProgress(response, 'upload');
                pollUploadJob(activeUploadJobId);
            },
            error: function (xhr, status, error) {
                uploadButton.disabled = false;
                startRequestInFlight = false;
                renderJobFailure('The start response was not received. Retry the same file to reuse the same request ID, or wait while the page searches for the server job. ' + (xhr.responseText || error));
                discoverMissingJobs();
            }
        });
    }

    function isBlockingStatus(status) {
        return ['Queued', 'ReceivingUpload', 'Validating', 'ReadyToSave', 'QueuedForSave', 'Saving',
            'PreparingSap', 'SendingToSap', 'SavingApproval', 'ValidationFailed', 'ReconciliationRequired',
            'Completed', 'Failed'].indexOf(status) >= 0;
    }

    function ensureNoBlockingJob(actionText) {
        if (!jobDiscoveryComplete) {
            alert('The page is still checking for an existing server job. Please try again in a moment.');
            return false;
        }
        const uploadUnknown = activeUploadJobId && !activeUploadStatus;
        const approvalUnknown = activeApprovalJobId && !activeApprovalStatus;
        if (startRequestInFlight || uploadUnknown || approvalUnknown || isBlockingStatus(activeUploadStatus) || isBlockingStatus(activeApprovalStatus)) {
            const approvalBlocks = approvalUnknown || isBlockingStatus(activeApprovalStatus);
            const id = approvalBlocks ? activeApprovalJobId : activeUploadJobId;
            alert('An upload or approval job is already in progress' + (id ? ' (Job ID: ' + id + ')' : '') + '. Open its progress before you ' + actionText + '.');
            if (id) {
                showStoredJob(approvalBlocks ? 'approval' : 'upload');
            }
            return false;
        }
        return true;
    }

    function restoreJobsOnLoad() {
        const uploadState = readJobState(uploadStorageKey);
        const approvalState = readJobState(approvalStorageKey);
        if (uploadState && uploadState.jobId) {
            activeUploadJobId = uploadState.jobId;
            activeUploadStatus = uploadState.status;
        }
        if (approvalState && approvalState.jobId) {
            activeApprovalJobId = approvalState.jobId;
            activeApprovalStatus = approvalState.status;
        }

        let remaining = 0;
        function finishedDiscovery() {
            remaining -= 1;
            if (remaining > 0) return;
            jobDiscoveryComplete = true;
            chooseInitialJob(uploadState, approvalState);
        }

        if (!activeUploadJobId) {
            remaining += 1;
            discoverLatestUploadJob(finishedDiscovery);
        }
        if (!activeApprovalJobId) {
            remaining += 1;
            discoverLatestApprovalJob(finishedDiscovery);
        }
        if (remaining === 0) {
            jobDiscoveryComplete = true;
            chooseInitialJob(uploadState, approvalState);
        }
    }

    function discoverMissingJobs() {
        if (!activeUploadJobId) discoverLatestUploadJob(function () {});
        if (!activeApprovalJobId) discoverLatestApprovalJob(function () {});
    }

    function discoverLatestUploadJob(done) {
        const pendingState = readJobState(uploadStorageKey) || {};
        $.ajax({
            url: 'Handler/UploadHandler.ashx?action=getUploadJobStatus', type: 'GET', dataType: 'json',
            data: { clientRequestId: pendingState.jobId ? '' : (pendingState.clientRequestId || '') },
            cache: false,
            timeout: draftOtbRequestTimeoutMs,
            success: function (job) {
                if (job && job.success && job.jobId) registerDiscoveredJob('upload', job);
            },
            complete: done
        });
    }

    function discoverLatestApprovalJob(done) {
        const pendingState = readJobState(approvalStorageKey) || {};
        $.ajax({
            url: 'Handler/DataOTBHandler.ashx?action=latestApprovalJob', type: 'GET', dataType: 'json',
            data: { clientRequestId: pendingState.jobId ? '' : (pendingState.clientRequestId || '') },
            cache: false,
            timeout: draftOtbRequestTimeoutMs,
            success: function (job) {
                if (job && job.success && job.jobId) registerDiscoveredJob('approval', job);
            },
            complete: done
        });
    }

    function registerDiscoveredJob(jobType, job) {
        if (clearServerAcknowledgedJob(jobType, job)) return;
        const key = jobType === 'approval' ? approvalStorageKey : uploadStorageKey;
        const existing = readJobState(key) || {};
        existing.jobId = job.jobId;
        existing.status = job.status;
        existing.updatedAt = job.updatedAt || new Date().toISOString();
        writeJobState(key, existing);
        if (jobType === 'approval') {
            activeApprovalJobId = job.jobId;
            activeApprovalStatus = job.status;
        } else {
            activeUploadJobId = job.jobId;
            activeUploadStatus = job.status;
        }
        updateJobSwitcher();
        if (jobDiscoveryComplete) {
            displayedJobType = jobType;
            draftOtbJobModal.show();
            if (jobType === 'approval') pollApprovalJob(job.jobId);
            else pollUploadJob(job.jobId);
        }
    }

    function chooseInitialJob(uploadState, approvalState) {
        const uploadUpdated = Date.parse((readJobState(uploadStorageKey) || uploadState || {}).updatedAt || 0) || 0;
        const approvalUpdated = Date.parse((readJobState(approvalStorageKey) || approvalState || {}).updatedAt || 0) || 0;
        if (activeUploadJobId || activeApprovalJobId) {
            if (activeApprovalStatus === 'ReconciliationRequired') displayedJobType = 'approval';
            else if (activeUploadStatus === 'ReconciliationRequired') displayedJobType = 'upload';
            else displayedJobType = activeApprovalJobId && (!activeUploadJobId || approvalUpdated >= uploadUpdated) ? 'approval' : 'upload';
            const selectedStatus = displayedJobType === 'approval' ? activeApprovalStatus : activeUploadStatus;
            const selectedJobId = displayedJobType === 'approval' ? activeApprovalJobId : activeUploadJobId;
            renderPersistedJobIdentity(displayedJobType, selectedJobId, selectedStatus);
            if (selectedStatus === 'ReconciliationRequired') {
                currentReconciliation = {
                    jobType: displayedJobType,
                    jobId: displayedJobType === 'approval' ? activeApprovalJobId : activeUploadJobId
                };
                setReconciliationControls(true);
            }
            draftOtbJobModal.show();
            if (activeUploadJobId) pollUploadJob(activeUploadJobId);
            if (activeApprovalJobId) pollApprovalJob(activeApprovalJobId);
        }
        updateJobSwitcher();
    }

    function showStoredJob(jobType) {
        if (currentReconciliation && currentReconciliation.jobType !== jobType) {
            alert('A reconciliation-required job must be acknowledged before viewing another job. Job ID: ' + currentReconciliation.jobId);
            return;
        }
        displayedJobType = jobType;
        currentReconciliation = null;
        const status = jobType === 'approval' ? activeApprovalStatus : activeUploadStatus;
        const storedJobId = jobType === 'approval' ? activeApprovalJobId : activeUploadJobId;
        renderPersistedJobIdentity(jobType, storedJobId, status);
        if (status === 'ReconciliationRequired') {
            currentReconciliation = { jobType: jobType, jobId: jobType === 'approval' ? activeApprovalJobId : activeUploadJobId };
            setReconciliationControls(true);
        } else {
            setReconciliationControls(false);
        }
        draftOtbJobModal.show();
        if (jobType === 'approval' && activeApprovalJobId) pollApprovalJob(activeApprovalJobId);
        if (jobType === 'upload' && activeUploadJobId) pollUploadJob(activeUploadJobId);
        updateJobSwitcher();
    }

    function renderPersistedJobIdentity(jobType, jobId, status) {
        if (!jobId) return;
        const terminal = ['Completed', 'ValidationFailed', 'Failed', 'ReconciliationRequired'].indexOf(status) >= 0;
        renderJobProgress({
            jobId: jobId,
            status: status || 'Restoring',
            stage: status || 'Restoring progress',
            progress: terminal ? 100 : 0,
            processedRows: 0,
            totalRows: 0,
            successRows: 0,
            errorRows: 0,
            message: status === 'ReconciliationRequired'
                ? 'Restoring reconciliation details for Job ID ' + jobId + '. Do not retry this job.'
                : 'Restoring the latest server progress for Job ID ' + jobId + '.'
        }, jobType);
    }

    function updateJobSwitcher() {
        $('#btnViewUploadJob').toggleClass('d-none', !activeUploadJobId).toggleClass('active', displayedJobType === 'upload');
        $('#btnViewApprovalJob').toggleClass('d-none', !activeApprovalJobId).toggleClass('active', displayedJobType === 'approval');
        $('#draftOtbJobSwitcher').toggleClass('d-none', !(activeUploadJobId && activeApprovalJobId));
        $('#btnOpenUploadJob').toggleClass('d-none', !activeUploadJobId);
        $('#btnOpenApprovalJob').toggleClass('d-none', !activeApprovalJobId);
        $('#draftOtbActiveJobBar').toggleClass('d-none', !(activeUploadJobId || activeApprovalJobId));
    }

    function resetJobModal(title, jobType) {
        $('#draftOtbJobModalLabel').text(title || 'Draft OTB progress');
        $('#btnSaveValidatedUpload, #btnDownloadUploadErrors, #btnAcknowledgeValidation').addClass('d-none');
        $('#draftOtbReconciliationDetails').addClass('d-none');
        $('#draftOtbReconciliationResultTable').empty();
        $('#btnAcknowledgeReconciliation').prop('disabled', true);
        renderJobProgress({
            stage: 'Queued', progress: 0, processedRows: 0, totalRows: 0,
            successRows: 0, errorRows: 0,
            message: 'The server is preparing this job.'
        }, jobType || 'upload');
    }

    function renderJobProgress(job, jobType) {
        if (displayedJobType && displayedJobType !== jobType) {
            updateJobSwitcher();
            return;
        }
        displayedJobType = jobType;
        $('#btnSaveValidatedUpload, #btnDownloadUploadErrors, #btnAcknowledgeValidation').addClass('d-none');
        const percent = Math.max(0, Math.min(100, parseInt(job.progress || 0, 10)));
        const total = parseInt(job.totalRows || 0, 10);
        const processed = parseInt(job.processedRows || 0, 10);
        const success = parseInt(job.successRows || 0, 10);
        const errors = parseInt(job.errorRows || 0, 10);
        $('#draftOtbJobModalLabel').text(jobType === 'approval' ? 'Draft OTB approval progress' : 'Draft OTB upload progress');
        $('#draftOtbJobStage').text(job.stage || job.status || 'Processing');
        $('#draftOtbJobPercent').text(percent + '%');
        $('#draftOtbJobProgressBar').css('width', percent + '%').text(percent + '%').attr('aria-valuenow', percent);
        $('#draftOtbJobCounts').text('Processed ' + processed.toLocaleString() + ' / ' + total.toLocaleString() +
            ' rows | Success: ' + success.toLocaleString() + ' | Error: ' + errors.toLocaleString());
        $('#draftOtbJobMessage').removeClass('alert-danger alert-success alert-warning').addClass('alert-info')
            .text(job.message || 'Processing continues on the server.');
        const jobId = job.jobId || (jobType === 'approval' ? activeApprovalJobId : activeUploadJobId);
        $('#draftOtbJobIdentity').toggleClass('d-none', !jobId);
        $('#draftOtbJobId').text(jobId || '');
        $('#draftOtbCloseSafety').text(jobId
            ? 'You may close or refresh this page. Processing continues on the server and this progress will be restored when you return.'
            : 'Keep this page open until the server returns a Job ID. Closing during file transfer can cancel the upload.');
        updateJobSwitcher();
    }

    function renderJobFailure(message) {
        $('#draftOtbJobStage').text('Failed');
        $('#draftOtbJobMessage').removeClass('alert-info alert-success alert-warning').addClass('alert-danger')
            .text(message || 'The job failed.');
    }

    function rememberJobStatus(jobType, job) {
        const key = jobType === 'approval' ? approvalStorageKey : uploadStorageKey;
        const state = readJobState(key) || {};
        state.jobId = job.jobId || state.jobId;
        state.status = job.status;
        state.updatedAt = job.updatedAt || new Date().toISOString();
        writeJobState(key, state);
        if (jobType === 'approval') activeApprovalStatus = job.status;
        else activeUploadStatus = job.status;
    }

    function getPollingDelay(retryCount) {
        return Math.min(15000, 1500 * Math.pow(2, Math.min(retryCount || 0, 3)));
    }

    function showPollingWarning(jobType) {
        if (displayedJobType !== jobType) return;
        $('#draftOtbJobMessage').removeClass('alert-danger alert-success alert-info').addClass('alert-warning')
            .text('Progress is temporarily unavailable. The server job is still running; this page will keep retrying.');
    }

    function clearStoredJob(jobType, expectedJobId) {
        const isApproval = jobType === 'approval';
        const currentJobId = isApproval ? activeApprovalJobId : activeUploadJobId;
        if (expectedJobId && currentJobId && String(expectedJobId).toLowerCase() !== String(currentJobId).toLowerCase()) {
            return false;
        }
        localStorage.removeItem(isApproval ? approvalStorageKey : uploadStorageKey);
        if (isApproval) {
            activeApprovalJobId = null;
            activeApprovalStatus = null;
        } else {
            activeUploadJobId = null;
            activeUploadStatus = null;
        }
        updateJobSwitcher();
        return true;
    }

    function clearServerAcknowledgedJob(jobType, job) {
        if (!job || job.acknowledged !== true) return false;
        if (currentReconciliation && currentReconciliation.jobType === jobType &&
            String(currentReconciliation.jobId).toLowerCase() === String(job.jobId || '').toLowerCase()) {
            currentReconciliation = null;
            setReconciliationControls(false);
        }
        clearStoredJob(jobType, job.jobId);
        if (displayedJobType === jobType) draftOtbJobModal.hide();
        return true;
    }

    function retainTerminalJobUntilProgressClosed(jobType, jobId) {
        $('#draftOtbJobModal').one('hidden.bs.modal', function () {
            clearStoredJob(jobType, jobId);
        });
        draftOtbJobModal.show();
    }

    function pollUploadJob(jobId, retryCount) {
        if (uploadPollTimer) clearTimeout(uploadPollTimer);
        $.ajax({
            url: 'Handler/UploadHandler.ashx?action=getUploadJobStatus',
            type: 'GET',
            dataType: 'json',
            data: { jobId: jobId, uploadBy: draftOtbCurrentUser },
            cache: false,
            timeout: draftOtbRequestTimeoutMs,
            success: function (job) {
                if (!job.success) {
                    if (job.retryable === false) {
                        renderJobFailure((job.message || 'This job is no longer available.') +
                            (job.errorCode === 'AUTH' ? ' Please sign in again, then reopen this page.' : ''));
                        return;
                    }
                    showPollingWarning('upload');
                    uploadPollTimer = setTimeout(function () { pollUploadJob(jobId, (retryCount || 0) + 1); }, getPollingDelay((retryCount || 0) + 1));
                    return;
                }
                if (clearServerAcknowledgedJob('upload', job)) return;
                rememberJobStatus('upload', job);
                renderJobProgress(job, 'upload');
                if (job.status === 'ReadyToSave') {
                    if (displayedJobType === 'upload') {
                        $('#draftOtbJobMessage').removeClass('alert-info alert-danger alert-warning').addClass('alert-success');
                        $('#btnSaveValidatedUpload').removeClass('d-none');
                    }
                    return;
                }
                if (job.status === 'ValidationFailed') {
                    if (displayedJobType === 'upload') {
                        $('#draftOtbJobMessage').removeClass('alert-info alert-success alert-warning').addClass('alert-danger');
                        if (job.hasErrorReport) {
                            $('#btnDownloadUploadErrors')
                                .attr('href', 'Handler/UploadHandler.ashx?action=downloadUploadErrors&jobId=' + encodeURIComponent(jobId) + '&uploadBy=' + encodeURIComponent(draftOtbCurrentUser))
                                .removeClass('d-none');
                        }
                        $('#btnAcknowledgeValidation').removeClass('d-none').text('I have downloaded/reviewed the validation errors');
                    }
                    // Retain the job so the full error report remains available after refresh.
                    return;
                }
                if (job.status === 'Completed') {
                    if (displayedJobType !== 'upload') return;
                    $('#draftOtbJobProgressBar').removeClass('progress-bar-animated');
                    $('#draftOtbJobMessage').removeClass('alert-info alert-danger alert-warning').addClass('alert-success');
                    $('#fileUpload').val('');
                    search();
                    retainTerminalJobUntilProgressClosed('upload', jobId);
                    return;
                }
                if (job.status === 'ReconciliationRequired') {
                    showReconciliationRequired('upload', job);
                    return;
                }
                if (job.status === 'Failed') {
                    if (displayedJobType !== 'upload') return;
                    renderJobFailure(job.message);
                    retainTerminalJobUntilProgressClosed('upload', jobId);
                    return;
                }
                uploadPollTimer = setTimeout(function () { pollUploadJob(jobId, 0); }, 1200);
            },
            error: function () {
                showPollingWarning('upload');
                uploadPollTimer = setTimeout(function () { pollUploadJob(jobId, (retryCount || 0) + 1); }, getPollingDelay((retryCount || 0) + 1));
            }
        });
    }

    function pollApprovalJob(jobId, retryCount) {
        if (approvalPollTimer) clearTimeout(approvalPollTimer);
        $.ajax({
            url: 'Handler/DataOTBHandler.ashx?action=approvaljobstatus',
            type: 'GET',
            dataType: 'json',
            data: { jobId: jobId, approvedBy: draftOtbCurrentUser },
            cache: false,
            timeout: draftOtbRequestTimeoutMs,
            success: function (job) {
                if (!job.success) {
                    if (job.retryable === false) {
                        renderJobFailure((job.message || 'This job is no longer available.') +
                            (job.errorCode === 'AUTH' ? ' Please sign in again, then reopen this page.' : ''));
                        return;
                    }
                    showPollingWarning('approval');
                    approvalPollTimer = setTimeout(function () { pollApprovalJob(jobId, (retryCount || 0) + 1); }, getPollingDelay((retryCount || 0) + 1));
                    return;
                }
                if (clearServerAcknowledgedJob('approval', job)) return;
                rememberJobStatus('approval', job);
                renderJobProgress(job, 'approval');
                const terminal = ['Completed', 'ValidationFailed', 'Failed', 'ReconciliationRequired'].indexOf(job.status) >= 0;
                if (!terminal) {
                    approvalPollTimer = setTimeout(function () { pollApprovalJob(jobId, 0); }, 1200);
                    return;
                }

                if (job.status === 'ReconciliationRequired') {
                    showReconciliationRequired('approval', job);
                    return;
                }

                if (displayedJobType !== 'approval') {
                    // Keep the terminal result available in the job switcher; do not
                    // interrupt another job that the user is currently monitoring.
                    return;
                }

                const succeeded = job.status === 'Completed';
                let usesResultModal = false;
                if (job.detailedResults && job.detailedResults.length) {
                    currentApprovalResults = {
                        jobId: job.jobId,
                        status: job.status,
                        message: job.message || '',
                        results: job.detailedResults
                    };
                    $('#btnDownloadApprovalResults').removeClass('d-none');
                    const resultSummary = buildApprovalResultTable(job.detailedResults, succeeded);
                    $('#approvalResultSummary').text((job.message || (succeeded ? 'Approval completed.' : 'Approval failed.')) +
                        ' Job ID: ' + job.jobId + '. ' + resultSummary)
                        .removeClass('alert-info alert-success alert-danger')
                        .addClass(succeeded ? 'alert-success' : 'alert-danger');
                    draftOtbJobModal.hide();
                    setTimeout(function () { approvalResultModal.show(); }, 250);
                    usesResultModal = true;
                    $('#approvalResultModal').one('hidden.bs.modal', function () {
                        search();
                        if (job.status === 'ValidationFailed') {
                            setTimeout(function () { draftOtbJobModal.show(); }, 250);
                        } else {
                            currentApprovalResults = null;
                            $('#btnDownloadApprovalResults').addClass('d-none');
                            clearStoredJob('approval', job.jobId);
                        }
                    });
                } else if (succeeded) {
                    currentApprovalResults = null;
                    $('#btnDownloadApprovalResults').addClass('d-none');
                    $('#draftOtbJobProgressBar').removeClass('progress-bar-animated');
                    $('#draftOtbJobMessage').removeClass('alert-info alert-danger alert-warning').addClass('alert-success');
                    search();
                } else {
                    currentApprovalResults = null;
                    $('#btnDownloadApprovalResults').addClass('d-none');
                    renderJobFailure(job.message);
                }
                if (job.status === 'ValidationFailed') {
                    // Preserve validation details and Job ID across refresh until a new approval is started.
                    $('#btnAcknowledgeValidation').removeClass('d-none').text('I have reviewed the validation result');
                    return;
                }
                if (!usesResultModal) {
                    retainTerminalJobUntilProgressClosed('approval', job.jobId);
                }
            },
            error: function () {
                showPollingWarning('approval');
                approvalPollTimer = setTimeout(function () { pollApprovalJob(jobId, (retryCount || 0) + 1); }, getPollingDelay((retryCount || 0) + 1));
            }
        });
    }

    function showReconciliationRequired(jobType, job) {
        if (currentReconciliation && currentReconciliation.jobType !== jobType) {
            // Keep the reconciliation record visible until it is acknowledged.
            // The second record remains persisted and will be presented next.
            updateJobSwitcher();
            return;
        }
        displayedJobType = jobType;
        currentReconciliation = {
            jobType: jobType,
            jobId: job.jobId,
            status: job.status,
            message: job.message || '',
            reviewed: false,
            detailedResults: Array.isArray(job.detailedResults) ? job.detailedResults : []
        };
        renderJobProgress(job, jobType);
        $('#draftOtbJobStage').text('MANUAL RECONCILIATION REQUIRED');
        $('#draftOtbJobMessage').removeClass('alert-info alert-success alert-warning').addClass('alert-danger')
            .text((job.message || 'The SAP outcome is uncertain.') + ' Do not retry. Record Job ID ' + job.jobId + ' and reconcile SAP against BMS.');
        setReconciliationControls(true);
        renderReconciliationReview(job);
        draftOtbJobModal.show();
    }

    function setReconciliationControls(required) {
        $('#btnAcknowledgeReconciliation').toggleClass('d-none', !required);
        if (required) {
            $('#btnAcknowledgeValidation').addClass('d-none');
            $('#btnAcknowledgeReconciliation').prop('disabled', true);
        } else {
            $('#draftOtbReconciliationDetails').addClass('d-none');
            $('#draftOtbReconciliationResultTable').empty();
        }
        $('.job-modal-close').toggleClass('d-none', required);
        $('#draftOtbJobModal .modal-header .btn-close').toggleClass('d-none', required);
    }

    function renderReconciliationReview(job) {
        const results = currentReconciliation ? currentReconciliation.detailedResults : [];
        $('#draftOtbReconciliationDetails').removeClass('d-none');
        $('#btnAcknowledgeReconciliation').prop('disabled', true);
        $('#btnReviewReconciliationResults').prop('disabled', false);
        if (results.length) {
            const summary = buildApprovalResultTable(results, false, 'draftOtbReconciliationResultTable');
            $('#draftOtbReconciliationSummary').text('Review the SAP response before acknowledging this reconciliation job. ' + summary);
            $('#btnReviewReconciliationResults').text('I have reviewed these results');
            $('#btnDownloadReconciliationResults').removeClass('d-none');
            $('#draftOtbReconciliationResultTable').removeClass('d-none');
        } else {
            $('#draftOtbReconciliationSummary').text('No row-level response was returned. Review the message and record Job ID ' + job.jobId + ' before acknowledging.');
            $('#btnReviewReconciliationResults').text('I reviewed the message and Job ID');
            $('#btnDownloadReconciliationResults').addClass('d-none');
            $('#draftOtbReconciliationResultTable').empty().addClass('d-none');
        }
    }

    function markReconciliationReviewed() {
        if (!currentReconciliation) return;
        currentReconciliation.reviewed = true;
        $('#btnReviewReconciliationResults').text('Reviewed').prop('disabled', true);
        $('#btnAcknowledgeReconciliation').prop('disabled', false);
    }

    function downloadReconciliationResults() {
        if (!currentReconciliation || !currentReconciliation.detailedResults.length) return;
        const downloadPayload = {
            jobId: currentReconciliation.jobId,
            jobType: currentReconciliation.jobType,
            status: currentReconciliation.status,
            message: currentReconciliation.message,
            downloadedAt: new Date().toISOString(),
            results: currentReconciliation.detailedResults
        };
        downloadJsonResultFile(downloadPayload, 'Draft_OTB_reconciliation_' + currentReconciliation.jobId + '.json');
        markReconciliationReviewed();
    }

    function downloadApprovalResults() {
        if (!currentApprovalResults || !currentApprovalResults.results || !currentApprovalResults.results.length) return;
        const downloadPayload = {
            jobId: currentApprovalResults.jobId,
            jobType: 'approval',
            status: currentApprovalResults.status,
            message: currentApprovalResults.message,
            downloadedAt: new Date().toISOString(),
            results: currentApprovalResults.results
        };
        downloadJsonResultFile(downloadPayload, 'Draft_OTB_approval_results_' + currentApprovalResults.jobId + '.json');
    }

    function downloadJsonResultFile(downloadPayload, fileName) {
        const blob = new Blob([JSON.stringify(downloadPayload, null, 2)], { type: 'application/json;charset=utf-8' });
        const objectUrl = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = objectUrl;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(function () { URL.revokeObjectURL(objectUrl); }, 1000);
    }

    function sendJobAcknowledgement(jobType, jobId, button, onSuccess) {
        const url = jobType === 'approval'
            ? 'Handler/DataOTBHandler.ashx?action=acknowledgeApprovalJob'
            : 'Handler/UploadHandler.ashx?action=acknowledgeUploadJob';
        button.disabled = true;
        $.ajax({
            url: url,
            type: 'POST',
            dataType: 'json',
            data: { jobId: jobId },
            timeout: draftOtbRequestTimeoutMs,
            success: function (response) {
                if (!response || !response.success) {
                    button.disabled = false;
                    alert((response && response.message) || 'The server could not acknowledge this job. Please try again.');
                    return;
                }
                onSuccess();
            },
            error: function (xhr, status, error) {
                button.disabled = false;
                alert('The acknowledgement was not saved. Please try again. ' + (xhr.responseText || error));
            }
        });
    }

    function acknowledgeValidationFailure() {
        const type = displayedJobType;
        const status = type === 'approval' ? activeApprovalStatus : activeUploadStatus;
        if (status !== 'ValidationFailed') return;
        const jobId = type === 'approval' ? activeApprovalJobId : activeUploadJobId;
        if (!jobId) return;
        if (!confirm('Confirm that you reviewed the validation result. The acknowledgement will be saved on the server.')) return;
        const button = this;
        sendJobAcknowledgement(type, jobId, button, function () {
            clearStoredJob(type, jobId);
            if (type === 'approval') {
                currentApprovalResults = null;
                $('#btnDownloadApprovalResults').addClass('d-none');
            }
            $('#btnAcknowledgeValidation').addClass('d-none');
            if (type !== 'approval' && activeApprovalStatus === 'ValidationFailed') showStoredJob('approval');
            else if (type !== 'upload' && activeUploadStatus === 'ValidationFailed') showStoredJob('upload');
            else draftOtbJobModal.hide();
        });
    }

    function acknowledgeReconciliation() {
        if (!currentReconciliation) return;
        const type = currentReconciliation.jobType;
        const jobId = currentReconciliation.jobId;
        if (!currentReconciliation.reviewed) {
            alert('Review the returned results, or the message and Job ID, before acknowledging this job.');
            return;
        }
        if (!confirm('Confirm that you recorded the Job ID and will reconcile SAP against BMS before any retry.')) return;
        const button = this;
        sendJobAcknowledgement(type, jobId, button, function () {
            clearStoredJob(type, jobId);
            currentReconciliation = null;
            setReconciliationControls(false);
            if (type !== 'approval' && activeApprovalStatus === 'ReconciliationRequired') showStoredJob('approval');
            else if (type !== 'upload' && activeUploadStatus === 'ReconciliationRequired') showStoredJob('upload');
            else draftOtbJobModal.hide();
        });
    }

    function copyDisplayedJobId() {
        const value = $('#draftOtbJobId').text();
        if (!value) return;
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(value);
        } else {
            const input = $('<textarea>').val(value).appendTo(document.body).select();
            document.execCommand('copy');
            input.remove();
        }
        $('#btnCopyDraftOtbJobId').text('Copied');
        setTimeout(function () { $('#btnCopyDraftOtbJobId').text('Copy Job ID'); }, 1200);
    }

    function buildApprovalPreviewTable(groups, totalRows) {
        const rows = groups || [];
        const displayedRows = rows.slice(0, approvalPreviewDisplayLimit);
        const omittedGroups = Math.max(0, rows.length - displayedRows.length);
        const html = [];
        html.push('<div class="mb-2"><strong>Selected rows:</strong> ' + Number(totalRows || 0).toLocaleString() + '</div>');
        html.push('<div class="mb-2"><strong>Budget groups:</strong> ' + rows.length.toLocaleString() + '</div>');
        if (omittedGroups > 0) {
            html.push('<div class="alert alert-warning py-2">Showing the first ' + displayedRows.length.toLocaleString() +
                ' groups. ' + omittedGroups.toLocaleString() + ' additional groups are included in the confirmed server snapshot but omitted here to keep this page responsive.</div>');
        }
        html.push('<table class="table table-bordered table-striped table-sm">');
        html.push('<thead class="table-primary"><tr><th>Company</th><th>Year</th><th>Month</th><th>Category</th>');
        html.push('<th class="text-end">Revised</th><th class="text-end">Diff</th><th class="text-end">Total Budget</th></tr></thead><tbody>');
        displayedRows.forEach(function (row) {
            html.push('<tr><td>' + htmlEncode(row.Company) + '</td><td>' + htmlEncode(row.Year) + '</td><td>' + htmlEncode(row.Month) +
                '</td><td>' + htmlEncode(row.Category) + '</td><td class="text-end">' + formatMoney(row.Revised) +
                '</td><td class="text-end">' + formatMoney(row.Diff) + '</td><td class="text-end fw-bold">' + formatMoney(row.TotalBudget) + '</td></tr>');
        });
        html.push('</tbody></table>');
        $('#approvalPreviewTableContainer').html(html.join(''));
    }

    function startConfirmedApprovalJob() {
        if (!pendingApprovalRunNos.length) return;
        if (!pendingApprovalPreviewHash) {
            alert('The approval preview token is missing. Close this dialog and prepare the preview again.');
            return;
        }
        if (!ensureNoBlockingJob('start another approval')) return;
        const button = this;
        button.disabled = true;
        const approvalSignature = pendingApprovalPreviewHash + '|' + pendingApprovalRunNos.slice().sort(function (a, b) { return Number(a) - Number(b); }).join(',');
        const requestState = getOrCreateRequestState(approvalStorageKey, 'approval', approvalSignature);
        const formData = new FormData();
        formData.append('runNos', JSON.stringify(pendingApprovalRunNos));
        formData.append('approvedBy', draftOtbCurrentUser);
        formData.append('previewHash', pendingApprovalPreviewHash);
        formData.append('clientRequestId', requestState.clientRequestId);
        startRequestInFlight = true;
        $.ajax({
            url: 'Handler/DataOTBHandler.ashx?action=approveDraftOTB',
            type: 'POST', data: formData, processData: false, contentType: false, dataType: 'json',
            timeout: draftOtbRequestTimeoutMs,
            success: function (job) {
                button.disabled = false;
                startRequestInFlight = false;
                if (!job.success) {
                    alert(job.message || 'Unable to start approval.');
                    return;
                }
                approvalPreviewModal.hide();
                resetJobModal('Draft OTB approval progress', 'approval');
                activeApprovalJobId = job.jobId;
                activeApprovalStatus = job.status || 'Queued';
                displayedJobType = 'approval';
                requestState.jobId = activeApprovalJobId;
                requestState.status = activeApprovalStatus;
                requestState.updatedAt = job.updatedAt || new Date().toISOString();
                writeJobState(approvalStorageKey, requestState);
                renderJobProgress(job, 'approval');
                setTimeout(function () { draftOtbJobModal.show(); }, 250);
                pollApprovalJob(activeApprovalJobId);
            },
            error: function (xhr, status, error) {
                button.disabled = false;
                startRequestInFlight = false;
                alert('The start response was not received. Confirm again to reuse the same request ID, or wait while the page searches for the server job. ' + (xhr.responseText || error));
                discoverMissingJobs();
            }
        });
    }

    function formatMoney(value) {
        const number = Number(value || 0);
        return number.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    function htmlEncode(value) {
        return $('<div>').text(value == null ? '' : value).html();
    }

    let showLoading = function (show = true, message = 'กำลังโหลดข้อมูล...', subMessage = 'กรุณารอสักครู่') {
        const overlay = document.getElementById('loadingOverlay');
        const loadingText = overlay.querySelector('.loading-text');
        const loadingSubtext = overlay.querySelector('.loading-subtext');

        if (show) {
            loadingText.textContent = message;
            loadingSubtext.textContent = subMessage;
            overlay.classList.add('active');
            document.body.style.overflow = 'hidden'; // ป้องกันการ scroll
        } else {
            overlay.classList.remove('active');
            document.body.style.overflow = ''; // คืนค่าการ scroll
        }
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
            
            // Here you would implement page content loading
            // Example: Load different content based on pageName
            if (pageName === 'Draft OTB Plan') {
                // Load Draft OTB Plan content
                loadDraftOTBContent();
            } else if (pageName === 'Approved OTB Plan') {
                // Load Approved OTB Plan content
                loadApprovedOTBContent();
            }
            // Add more conditions for other pages...
        }

        // Example function to load Draft OTB content
        function loadDraftOTBContent() {
            console.log('Loading Draft OTB Plan...');
            // Implementation for loading Draft OTB content
        }

        // Example function to load Approved OTB content
        function loadApprovedOTBContent() {
            console.log('Loading Approved OTB Plan...');
            // Implementation for loading Approved OTB content
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
        let initial = function () {
            const firstMenuLink = document.querySelector('.menu-link');
            if (firstMenuLink) {
                firstMenuLink.classList.add('expanded');
            }

            if (typeDropdown) {
                $(typeDropdown).select2({
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

            $('#DDSegment').on('select2:select', changeVendor);
            btnClearFilter.addEventListener('click', function () {
                mainForm.reset();

                $("#DDType").val(null).trigger('change');
                $("#DDYear").val(null).trigger('change');
                $("#DDMonth").val(null).trigger('change');
                $("#DDCompany").val(null).trigger('change');
                $("#DDSegment").val(null).trigger('change');
                $("#DDCategory").val(null).trigger('change');
                $("#DDBrand").val(null).trigger('change');
                $("#DDVendor").val(null).trigger('change');

                InitMSData();

                tableViewBody.innerHTML = "";
            });
            btnView.addEventListener('click', search);
            btnExportTXN.addEventListener('click', exportTXN);
            btnExportSUM.addEventListener('click', exportSum);

            // *** ADDED: Approve Button Click Event ***
            btnApprove.addEventListener('click', approveSelectedItems);
            btnDelete.addEventListener('click', deleteSelectedItems);
    }

    let deleteDraftOTB = function (runNo) {
        if (!confirm('Are you sure you want to delete this Draft OTB?')) {
            return;
        }
        var formData = new FormData();
        formData.append('runNo', runNo);
        $.ajax({
            url: 'Handler/DataOTBHandler.ashx?action=deleteDraftOTB',
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            success: function (response) {
                if (response.trim() === "Success") {
                    alert('Draft OTB deleted successfully!');
                    search(); // โหลดข้อมูลตารางใหม่
                } else {
                    alert('Error deleting Draft OTB: ' + response);
                }
            },
            error: function (xhr, status, error) {
                console.log('Error deleting Draft OTB: ' + error);
                alert('An error occurred while deleting the Draft OTB.');
            }
        });
    }

    // *** MODIFIED: Function to Approve Selected Items ***
    let approveSelectedItems = function () {
        const runNosToApprove = [];
        $('input[name="checkselect"]:checked').each(function () {
            const runNo = this.id.replace('checkselect', '');
            runNosToApprove.push(runNo);
        });

        if (runNosToApprove.length === 0) {
            alert('Please select items to approve.');
            return;
        }
        if (runNosToApprove.length > 15000) {
            alert('A maximum of 15,000 rows can be approved in one job.');
            return;
        }

        const formData = new FormData();
        formData.append('runNos', JSON.stringify(runNosToApprove));
        formData.append('approvedBy', draftOtbCurrentUser);
        showLoading(true, 'Preparing approval preview...', 'Calculating Company / Year / Month / Category totals');
        $.ajax({
            url: 'Handler/DataOTBHandler.ashx?action=approvalPreview',
            type: 'POST', data: formData, processData: false, contentType: false, dataType: 'json',
            success: function (response) {
                showLoading(false);
                if (!response.success) {
                    alert(response.message || 'Unable to prepare approval preview.');
                    return;
                }
                pendingApprovalRunNos = runNosToApprove;
                pendingApprovalPreviewHash = response.previewHash || '';
                if (!pendingApprovalPreviewHash) {
                    alert('The server did not return an approval preview token. Please refresh and try again.');
                    return;
                }
                buildApprovalPreviewTable(response.groups, response.totalRows);
                approvalPreviewModal.show();
            },
            error: function (xhr, status, error) {
                showLoading(false);
                alert('Unable to prepare approval preview: ' + (xhr.responseText || error));
            }
        });
    }

    // *** NEW: Function to build the approval result table ***
    function buildApprovalResultTable(results, isSuccess, containerId) {
        var container = document.getElementById(containerId || 'approvalResultTableContainer');
        if (!results || results.length === 0) {
            container.innerHTML = "<p>No result details were returned.</p>";
            return '';
        }

        const failedRows = results.filter(function (row) {
            return String(row.SAP_MessageType || 'E').toUpperCase() !== 'S';
        });
        const successfulRows = results.filter(function (row) {
            return String(row.SAP_MessageType || 'E').toUpperCase() === 'S';
        });
        const sourceRows = failedRows.concat(successfulRows);
        const displayedRows = sourceRows.slice(0, 500);
        var sb = [];
        if (sourceRows.length > displayedRows.length) {
            sb.push("<div class='alert alert-warning py-2'>Showing the first " + displayedRows.length.toLocaleString() +
                " of " + sourceRows.length.toLocaleString() + " result rows. Error rows are listed first.</div>");
        } else if (failedRows.length) {
            sb.push("<div class='alert alert-danger py-2'>Showing all results with " + failedRows.length.toLocaleString() + " error rows listed first.</div>");
        }
        sb.push("<table class='table table-bordered table-striped table-sm table-hover'>");
        sb.push("<thead class='table-primary sticky-header'><tr>");
        sb.push("<th>Year</th><th>Month</th><th>Category</th><th>Category Name</th>");
        sb.push("<th>Segment</th><th>Segment Name</th><th>Brand</th><th>Brand Name</th>");
        sb.push("<th>Vendor</th><th>Vendor Name</th><th class='text-end'>Amount (THB)</th>");
        sb.push("<th class='text-danger'>SAP Status</th><th class='text-danger'>SAP Message</th>");
        sb.push("</tr></thead><tbody>");

        displayedRows.forEach(function (row) {
            let statusType = String(row.SAP_MessageType || 'E').toUpperCase();
            let rowClass = (statusType === 'S') ? 'table-success' : 'table-danger';

            sb.push(`<tr class="${rowClass}">`);
            sb.push(`<td>${htmlEncode(row.OTBYear)}</td>`);
            sb.push(`<td>${htmlEncode(row.OTBMonth)}</td>`);
            sb.push(`<td>${htmlEncode(row.OTBCategory)}</td>`);
            sb.push(`<td>${htmlEncode(row.CateName)}</td>`);
            sb.push(`<td>${htmlEncode(row.OTBSegment)}</td>`);
            sb.push(`<td>${htmlEncode(row.SegmentName)}</td>`);
            sb.push(`<td>${htmlEncode(row.OTBBrand)}</td>`);
            sb.push(`<td>${htmlEncode(row.BrandName)}</td>`);
            sb.push(`<td>${htmlEncode(row.OTBVendor)}</td>`);
            sb.push(`<td>${htmlEncode(row.Vendor)}</td>`);
            sb.push(`<td class="text-end">${htmlEncode(formatMoney(row.Amount))}</td>`);
            sb.push(`<td><strong>${htmlEncode(statusType)}</strong></td>`);
            sb.push(`<td>${htmlEncode(row.SAP_Message || (statusType === 'S' ? 'Success' : 'No Message'))}</td>`);
            sb.push("</tr>");
        });

        sb.push("</tbody></table>");
        container.innerHTML = sb.join('');
        return 'Total results: ' + results.length.toLocaleString() + '; errors: ' + failedRows.length.toLocaleString() +
            '; displayed: ' + displayedRows.length.toLocaleString() + '.';
    }

    let deleteSelectedItems = function () {
        let runNosToDelete = [];

        $('input[name="checkselect"]:checked').each(function () {

            let runNo = this.id.replace('checkselect', '');
            runNosToDelete.push(runNo);
        });

        if (runNosToDelete.length === 0) {
            alert('Please select items to approve.');
            return;
        }

        if (!confirm('Are you sure you want to delete ' + runNosToDelete.length + ' selected items?')) {
            return;
        }

        var currentUser = '<%= HttpUtility.JavaScriptStringEncode(Session("user").ToString()) %>';
        var approvedBy = currentUser || 'unknown';

        var formData = new FormData();
        formData.append('runNos', JSON.stringify(runNosToDelete)); // ส่งเป็น JSON String
        formData.append('approvedBy', approvedBy);

        $.ajax({
            url: 'Handler/DataOTBHandler.ashx?action=deleteDraftOTB',
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            dataType: 'json',
            success: function (response) {
                if (response.success === true) {
                    alert('Items deleted successfully!');
                    tableViewBody.innerHTML = "";
                    search(); // โหลดข้อมูลตารางใหม่
                } else {
                    alert('Error deleting items.');
                }
            },
            error: function (xhr, status, error) {
                console.log('Error deleting items: ' + error);
                alert('An error occurred while deleting items.');
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

        showLoading(true, 'กำลังค้นหาข้อมูล...', 'กรุณารอสักครู่');

        $.ajax({
            url: 'Handler/DataOTBHandler.ashx?action=obtlistbyfilter',
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            success: function (response) {
                tableViewBody.innerHTML = response;
                showLoading(false);
            },
            error: function (xhr, status, error) {
                console.log('Error getlist data: ' + error);
                showLoading(false);
                alert('เกิดข้อผิดพลาดในการโหลดข้อมูล');
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

        BMSLoading.download('Handler/DataOTBHandler.ashx?' + params.toString(), 'Exporting Draft OTB...', 'Preparing Excel file');
    }

    let exportSum = function () {
        console.log("Export Sum clicked");

        // **สำคัญ**: เราจะใช้แค่ Filter 3 ตัวตามที่รูปภาพระบุ (Year, Company, Segment)
        // แม้ว่าหน้าเว็บจะมี Filter อื่นๆ ก็ตาม
        let OTBtype = typeDropdown.value; // (ดูเหมือน SP จะไม่ใช้ แต่ส่งไปเผื่อ)
        let OTByear = yearDropdown.value;
        let OTBmonth = monthDropdown.value; // (SP ไม่ใช้)
        let OTBcompany = companyDropdown.value;
        let OTBCategory = categoryDropdown.value; // (SP ไม่ใช้)
        let OTBSegment = segmentDropdown.value;
        let OTBBrand = brandDropdown.value; // (SP ไม่ใช้)
        let OTBVendor = vendorDropdown.value; // (SP ไม่ใช้)

        if (!OTByear) {
            alert("Please select a Year to export the summary.");
            return;
        }

        var params = new URLSearchParams();
        params.append('action', 'exportdraftotbsum'); // Action ใหม่
        params.append('OTByear', OTByear);

        // ส่งค่า Company และ Segment ถ้ามี
        if (OTBcompany) {
            params.append('OTBCompany', OTBcompany);
        }
        if (OTBSegment) {
            params.append('OTBSegment', OTBSegment);
        }

        BMSLoading.download('Handler/DataOTBHandler.ashx?' + params.toString(), 'Exporting Draft OTB Summary...', 'Preparing Excel file');
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
</html>
