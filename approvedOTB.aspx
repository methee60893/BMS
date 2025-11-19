<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="approvedOTB.aspx.vb" Inherits="BMS.approvedOTB" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - Approved OTB Plan</title>
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
                    <li><a href="approvedOTB.aspx" class="menu-link active">Approved OTB Plan</a></li>
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
                <button class="menu-toggle" onclick="toggleSidebar()">
                    <i class="bi bi-list"></i>
                </button>
                <h1 class="page-title" id="pageTitle">KBMS - Approved OTB</h1>
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
            <!-- Filter Box -->
            <div class="filter-box">
                <div class="filter-header">
                    <i class="bi bi-funnel"></i>
                    Approved OTB
                </div>
                <div class="filter-body">
                    <!-- Filter Fields -->
                    <div class="row g-3 mb-3">
                        <div class="col-md-3">
                            <label class="form-label">Type</label>
                            <select  id="DDType" class="form-select">
                                <option value=''>-- กรุณาเลือก Type --</option>
                                <option value="Original" >Original</option> 
                                <option value="Revise" >Revise</option>
                            </select>
                        </div>
                        <div class="col-md-3">
                            <label class="form-label">Version</label>
                            <select id="DDVersion" class="form-select"></select>
                        </div>
                        <div class="col-md-6">
                            <label class="form-label">Category</label>
                            <select id="DDCategory" class="form-select"></select>
                        </div>
                    </div>

                    <div class="row g-3 mb-3">
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
                        <div class="col-md-6">
                            <label class="form-label">Segment</label>
                            <select id="DDSegment" class="form-select">
                            </select>
                        </div>
                    </div>

                    <div class="row g-3 mb-4">
                        <div class="col-md-3">
                            <label class="form-label">Company</label>
                            <select id="DDCompany" class="form-select">

                            </select>
                        </div>
                        <div class="col-md-3">
                            <label class="form-label">Brand</label>
                            <select id="DDBrand" class="form-select"></select>
                        </div>
                        <div class="col-md-6">
                            <label class="form-label">Vendor</label>
                            <select id="DDVendor" class="form-select"></select>
                        </div>
                    </div>

                    <!-- Action Buttons -->
                    <div class="row">
                        <div class="col-12 text-end">
                            <button class="btn btn-clear btn-custom me-2" id="btnClearFilter">
                                <i class="bi bi-x-circle"></i> Clear Filter
                            </button>
                            <button class="btn btn-view btn-custom" id="btnView" >
                                <i class="bi bi-eye"></i> View
                            </button>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Export Button -->
            <div class="export-section">
                <!-- *** MODIFIED: Changed ID from btnExport to btnExportTXN for consistency *** -->
                <button class="btn btn-export btn-custom" id="btnExportTXN">
                    <i class="bi bi-file-earmark-excel"></i> Export TXN
                </button>
            </div>

            <!-- Data Table -->
            <div class="table-container">
                <div class="table-responsive">
                    <table class="table table-hover mb-0">
                        <thead>
                            <tr>
                                <th>Create Date</th>
                                <th>Version</th>
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
                                <th>Amount (THB)</th>
                                <th>Revised Diff</th>
                                <th>Remark</th>
                                <th>Status</th>
                                <th>Approved Date</th>
                                <th>Create By</th>
                                <th>Approved By</th>
                                <th>SAP Status</th>
                            </tr>
                        </thead>
                        <tbody id="tableViewBody">
                           
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script>

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
        let versionDropdown = document.getElementById("DDVersion");

        // *** ADDED: btnExportTXN variable ***
        let btnExportTXN = document.getElementById("btnExportTXN");


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
        document.addEventListener('click', function (event) {
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
            $('#DDType').on('select2:select', changeType);
            btnClearFilter.addEventListener('click', function () {

                // Reset ค่าใน Dropdown ทุกตัวด้วยตนเอง
                $("#DDType").val(null).trigger('change');
                $("#DDYear").val(null).trigger('change');
                $("#DDMonth").val(null).trigger('change');
                $("#DDCompany").val(null).trigger('change');
                $("#DDSegment").val(null).trigger('change');
                $("#DDCategory").val(null).trigger('change');
                $("#DDBrand").val(null).trigger('change');
                $("#DDVendor").val(null).trigger('change');
                $("#DDVersion").val(null).trigger('change');

               
                tableViewBody.innerHTML = "<tr><td colspan='24' class='text-center text-muted'>Filters cleared. Click View to search.</td></tr>";
            });
            btnView.addEventListener('click', search);

            // *** ADDED: Event listener for Export button ***
            btnExportTXN.addEventListener('click', exportTXN);

        }

        let search = function () {
            let segmentCode = segmentDropdown.value;
            let cate = categoryDropdown.value;
            let brandCode = brandDropdown.value;
            let vendorCode = vendorDropdown.value;
            let OTBtype = typeDropdown.value;
            let OTByear = yearDropdown.value;
            let OTBmonth = monthDropdown.value;
            let OTBcompany = companyDropdown.value;
            let OTBVersion = versionDropdown.value;

            var formData = new FormData();
            formData.append('OTBtype', OTBtype);
            formData.append('OTByear', OTByear);
            formData.append('OTBmonth', OTBmonth);
            formData.append('OTBCompany', OTBcompany);
            formData.append('OTBCategory', cate);
            formData.append('OTBSegment', segmentCode);
            formData.append('OTBBrand', brandCode);
            formData.append('OTBVendor', vendorCode);
            formData.append('OTBVersion', OTBVersion);

            $.ajax({
                url: 'Handler/DataOTBHandler.ashx?action=obtApprovelistbyfilter',
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

        // *** ADDED: exportTXN function ***
        let exportTXN = function () {
            console.log("Export TXN clicked");
            // Build query string from filters
            var params = new URLSearchParams();
            params.append('action', 'exportapprovedotb'); // <-- NEW ACTION
            params.append('OTBtype', typeDropdown.value);
            params.append('OTByear', yearDropdown.value);
            params.append('OTBmonth', monthDropdown.value);
            params.append('OTBCompany', companyDropdown.value);
            params.append('OTBCategory', categoryDropdown.value);
            params.append('OTBSegment', segmentDropdown.value);
            params.append('OTBBrand', brandDropdown.value);
            params.append('OTBVendor', vendorDropdown.value);
            params.append('OTBVersion', versionDropdown.value); // <-- Added Version

            // Use window.location to trigger file download (GET request)
            window.location.href = 'Handler/DataOTBHandler.ashx?' + params.toString();
        }

        let InitMSData = function () {
            InitSegment(segmentDropdown);
            InitCategoty(categoryDropdown);
            InitBrand(brandDropdown);
            InitVendor(vendorDropdown);
            InitMSYear(yearDropdown);
            InitMonth(monthDropdown);
            InitCompany(companyDropdown);
            InitVersion(versionDropdown);
          
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
        let InitVersion = function (versionDropdown) {
            var OTBtype = typeDropdown.value;

            var formData = new FormData();
            formData.append('OTBtype', OTBtype);
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=VersionMSList',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    versionDropdown.innerHTML = response;
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

        let changeType = function () {
            InitVersion(versionDropdown)
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
</body>
</html>