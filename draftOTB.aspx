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
            <button class="close-sidebar" type="button" onclick="toggleSidebar()">
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
                <h1 class="page-title" id="pageTitle">KBMS - Draft OTB</h1>
            </div>
            <div class="user-info">
                <span class="d-none d-md-inline">Welcome, <%= HttpUtility.JavaScriptStringEncode(Session("user").ToString()) %></span>
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

    <div class="loading-overlay" id="loadingOverlay">
    <div class="loading-content">
        <div class="loading-spinner"></div>
        <p class="loading-text">กำลังโหลดข้อมูล...</p>
        <p class="loading-subtext">กรุณารอสักครู่</p>
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

    // *** NEW: Add modal instance variable ***
    let approvalResultModal;


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
            console.log("Save from Preview button clicked");

            var selectedRows = [];

            $('#previewTableContainer input[name="selectedRows"]:checked').each(function () {
                var cb = $(this);
                // ดึงข้อมูลจาก data-attributes ที่เราเก็บไว้
                var rowData = {
                    Type: cb.data('type'),
                    Year: cb.data('year'),
                    Month: cb.data('month'),
                    Category: cb.data('category'),
                    Company: cb.data('company'),
                    Segment: cb.data('segment'),
                    Brand: cb.data('brand'),
                    Vendor: cb.data('vendor'),
                    Amount: cb.data('amount'),
                    Remark: cb.data('remark')
                    // เราไม่จำเป็นต้องส่ง 'canUpdate' เพราะ Server จะ Validate ซ้ำอีกครั้ง
                };
                selectedRows.push(rowData);
            });

            if (selectedRows.length === 0) {
                alert('Please select at least one row to save.');

                return;
            }

            if (!confirm('Confirm to save ' + selectedRows.length + ' selected row(s) to database?')) return;


            var currentUser = '<%= HttpUtility.JavaScriptStringEncode(Session("user").ToString()) %>';
            var uploadBy = currentUser || 'unknown';
            console.log(uploadBy);


            var formData = new FormData();
            formData.append('uploadBy', uploadBy);
            // แปลง Array ของข้อมูลเป็น JSON String แล้วส่งไป
            formData.append('selectedData', JSON.stringify(selectedRows));

            $.ajax({
                url: 'Handler/UploadHandler.ashx?action=savePreview',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    if (response.includes("alert-danger")) {
                        alert('Error saving data: ' + $(response).text());
                    } else {
                        alert(response);
                        $('#previewModal').modal('hide');
                        $('#previewTableContainer').empty();
                        $('#fileUpload').val('');
                        search(); 
                    }
                },
                error: function (xhr, status, error) {
                    alert('Error saving data: ' + error);
                }
            });
        });

        // *** NEW: Instantiate the new modal ***
        approvalResultModal = new bootstrap.Modal(document.getElementById('approvalResultModal'));
    });

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
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }

            if (yearDropdown) {
                $(yearDropdown).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }

            if (monthDropdown) {
                $(monthDropdown).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }

            if (companyDropdown) {
                $(companyDropdown).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }

            if (segmentDropdown) {
                $(segmentDropdown).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }

            if (categoryDropdown) {
                $(categoryDropdown).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }

            if (brandDropdown) {
                $(brandDropdown).select2({
                    theme: "bootstrap-5",
                    allowClear: true
                });
            }

            if (vendorDropdown) {
                $(vendorDropdown).select2({
                    theme: "bootstrap-5",
                    allowClear: true
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
        let runNosToApprove = [];
        // ค้นหา Checkbox ที่ชื่อ 'checkselect' ที่ถูกเลือก
        $('input[name="checkselect"]:checked').each(function () {
         
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
            
            // *** NEW: Show loading overlay ***
            showLoading(true, 'Approving items...', 'Contacting SAP');

            $.ajax({
                url: 'Handler/DataOTBHandler.ashx?action=approveDraftOTB',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                dataType: 'json', // <-- Ensure jQuery parses the response as JSON
                success: function (response) {
                    showLoading(false); // Hide loading
                    // Handler จะส่ง success:true หรือ success:false กลับมาเสมอ
                    if (response.success === true) {
                        // *** กรณีสำเร็จทั้งหมด ***
                        if (response.action === 'preview') {
                            // แสดง Modal สรุปผล (กรณีสำเร็จ)
                            buildApprovalResultTable(response.detailedResults, true);
                            $('#approvalResultSummary').text(response.message);
                            $('#approvalResultSummary').removeClass('alert-info alert-success alert-danger').addClass('alert-success');
                            approvalResultModal.show();

                            // เมื่อปิด Modal ให้โหลดข้อมูลใหม่
                            $('#approvalResultModal').one('hidden.bs.modal', function () {
                                search(); // Reload main grid
                            });
                        } else {
                            // Fallback (เผื่อ Logic เก่า)
                            alert(response.message);
                            search();
                        }

                    } else {
                        // *** กรณี Error (success: false) ***

                        if (response.action === 'preview' && response.detailedResults) {
                            // *** กรณี SAP Error บางส่วน (Partial) ***
                            buildApprovalResultTable(response.detailedResults, false);
                            $('#approvalResultSummary').text(response.message);
                            $('#approvalResultSummary').removeClass('alert-info alert-success alert-danger').addClass('alert-danger');
                            approvalResultModal.show();

                        } else {
                            // *** กรณี Error ทั่วไป (เช่น SAP Error ที่เรา Throw มา) ***
                            // แสดง alert ด้วยข้อความ Error ที่ชัดเจนจาก Server
                            alert('Error: ' + (response.message || 'An unknown error occurred.'));
                        }
                    }
                },
                error: function (xhr, status, error) {
                    showLoading(false); // Hide loading
                    console.log('Error approving items: ' + error, xhr.responseText);
                    alert('An error occurred while approving items. ' + xhr.responseText);
                }
            });
    }

    // *** NEW: Function to build the approval result table ***
    function buildApprovalResultTable(results, isSuccess) {
        var container = document.getElementById('approvalResultTableContainer');
        if (!results || results.length === 0) {
            container.innerHTML = "<p>No result details were returned.</p>";
            return;
        }

        var sb = [];
        sb.push("<table class='table table-bordered table-striped table-sm table-hover'>");
        sb.push("<thead class='table-primary sticky-header'><tr>");
        sb.push("<th>Type</th><th>Year</th><th>Month</th><th>Category</th><th>Category Name</th>");
        sb.push("<th>Segment</th><th>Segment Name</th><th>Brand</th><th>Brand Name</th>");
        sb.push("<th>Vendor</th><th>Vendor Name</th><th class='text-end'>Amount (THB)</th>");
        sb.push("<th class='text-danger'>SAP Status</th><th class='text-danger'>SAP Message</th>");
        sb.push("</tr></thead><tbody>");

        results.forEach(row => {
            let statusType = (row.SAP_MessageType || 'E').toUpperCase();
            let rowClass = (statusType === 'S') ? 'table-success' : 'table-danger';
            
            sb.push(`<tr class="${rowClass}">`);
            sb.push(`<td>${row.OTBType || ''}</td>`);
            sb.push(`<td>${row.OTBYear || ''}</td>`);
            sb.push(`<td>${row.OTBMonth || ''}</td>`); // Assuming month number is fine
            sb.push(`<td>${row.OTBCategory || ''}</td>`);
            sb.push(`<td>${row.CateName || ''}</td>`);
            sb.push(`<td>${row.OTBSegment || ''}</td>`);
            sb.push(`<td>${row.SegmentName || ''}</td>`);
            sb.push(`<td>${row.OTBBrand || ''}</td>`);
            sb.push(`<td>${row.BrandName || ''}</td>`);
            sb.push(`<td>${row.OTBVendor || ''}</td>`);
            sb.push(`<td>${row.Vendor || ''}</td>`);
            sb.push(`<td class="text-end">${parseFloat(row.Amount || 0).toFixed(2)}</td>`);
            sb.push(`<td><strong>${statusType}</strong></td>`);
            sb.push(`<td>${row.SAP_Message || (statusType === 'S' ? 'Success' : 'No Message')}</td>`);
            sb.push("</tr>");
        });

        sb.push("</tbody></table>");
        container.innerHTML = sb.join('');
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

        // Use window.location to trigger file download
        // This is a GET request, so the handler must be adjusted to read from QueryString
        window.location.href = 'Handler/DataOTBHandler.ashx?' + params.toString();
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

        // ใช้ window.location เพื่อดาวน์โหลดไฟล์ (GET request)
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
        search();
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