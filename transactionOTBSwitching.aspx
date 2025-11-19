<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="transactionOTBSwitching.aspx.vb" Inherits="BMS.transactionOTBSwitching" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - Switching Transaction</title>
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
                    <li><a href="transactionOTBSwitching.aspx" class="menu-link active">Switching Transaction</a></li>
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
                <h1 class="page-title" id="pageTitle">KBMS - Switching Transaction</h1>
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
                Transaction OTB Switching
            </div>

            <!-- Filter Box -->
            <div class="filter-box">
                <div class="row g-3 mb-3">
                    <div class="col-md-3">
                        <label class="form-label">Type</label>
                        <select id="DDSwitchType" class="form-select">
                            <option value="">--กรุณาเลือก Type --</option>
                            <option value="D">Switch In-out</option>
                            <option value="G">Carry In-out</option>
                            <option value="I">Balance In-Out</option>
                            <option value="E">Extra</option>
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

                <div class="row g-3 mb-3">
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
                        <select id="DDVendor" class="form-select">

                        </select>
                    </div>
                </div>

                <!-- Action Buttons -->
                <div class="row">
                    <div class="col-12 text-end">
                        <button class="btn btn-clear btn-custom me-2" id="btnClearFilter">
                            <i class="bi bi-x-circle"></i> Clear Filter
                        </button>
                        <button class="btn btn-view btn-custom" id="btnView">
                            <i class="bi bi-eye"></i> View
                        </button>
                    </div>
                </div>
            </div>

            <!-- Export Button -->
            <div class="export-section">
                <button class="btn btn-export btn-custom" id="btnExport">
                    <i class="bi bi-file-earmark-excel"></i> Export TXN
                </button>
            </div>

            <!-- Data Table -->
            <div class="table-container">
                <div class="table-responsive">
                    <table class="table table-bordered mb-0">
                        <thead>
                            <tr>
                                <th rowspan="2">Create date</th>
                                <th rowspan="2">Type</th>
                                <th colspan="11" style="background: #5fa8d3;">out</th>
                                <th rowspan="2">Type</th>
                                <th colspan="11" style="background: #5fa8d3;">in</th>
                                <th rowspan="2">Amount (THB)</th>
                                <th rowspan="2">Create by</th>
                            </tr>
                            <tr>
                                <!-- Out columns -->
                                <th rowspan="1">Year</th>
                                <th rowspan="1">Month</th>
                                <th rowspan="1">Category</th>
                                <th rowspan="1">Category name</th>
                                <th rowspan="1">Company</th>
                                <th rowspan="1">Segment</th>
                                <th rowspan="1">Segment name</th>
                                <th rowspan="1">Brand</th>
                                <th rowspan="1">Brand name</th>
                                <th rowspan="1">Vendor</th>
                                <th rowspan="1">Vendor name</th>
                                <!-- In columns -->
                                <th rowspan="1">Year</th>
                                <th rowspan="1">Month</th>
                                <th rowspan="1">Category</th>
                                <th rowspan="1">Category name</th>
                                <th rowspan="1">Company</th>
                                <th rowspan="1">Segment</th>
                                <th rowspan="1">Segment name</th>
                                <th rowspan="1">Brand</th>
                                <th rowspan="1">Brand name</th>
                                <th rowspan="1">Vendor</th>
                                <th rowspan="1">Vendor name</th>
                            </tr>
                        </thead>
                        <tbody id="tableViewBody">
                            <tr><td colspan='30' class='text-center text-muted'>No switch OTB records found</td></tr>
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
        // *** MODIFIED: Declare variables here, but assign them inside initial() ***
        let typeDropdown, yearDropdown, monthDropdown, companyDropdown;
        let segmentDropdown, categoryDropdown, brandDropdown, vendorDropdown;
        let btnClearFilter, btnView, btnExport;
        let tableViewBody; // <-- Added missing variable

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

        let initial = function () {
            // *** MODIFIED: Assign variables inside initial() after DOM is loaded ***
            typeDropdown = document.getElementById("DDSwitchType");
            yearDropdown = document.getElementById("DDYear");
            monthDropdown = document.getElementById("DDMonth");
            companyDropdown = document.getElementById("DDCompany");
            segmentDropdown = document.getElementById("DDSegment");
            categoryDropdown = document.getElementById("DDCategory");
            brandDropdown = document.getElementById("DDBrand");
            vendorDropdown = document.getElementById("DDVendor");
            btnClearFilter = document.getElementById("btnClearFilter");
            btnView = document.getElementById("btnView");
            btnExport = document.getElementById("btnExport");
            tableViewBody = document.getElementById("tableViewBody"); // <-- Added missing assignment

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

            // *** FIX: Check if elements exist before adding listeners ***
            if (segmentDropdown) {
                segmentDropdown.addEventListener('change', changeVendor);
            }

                btnClearFilter.addEventListener('click', function () {
                    //mainForm.reset(); // No mainForm here
                    // Manual reset
                    // Reset ค่าใน Dropdown ทุกตัวด้วยตนเอง
                    $("#DDSwitchType").val(null).trigger('change');
                    $("#DDYear").val(null).trigger('change');
                    $("#DDMonth").val(null).trigger('change');
                    $("#DDCompany").val(null).trigger('change');
                    $("#DDSegment").val(null).trigger('change');
                    $("#DDCategory").val(null).trigger('change');
                    $("#DDBrand").val(null).trigger('change');
                    $("#DDVendor").val(null).trigger('change');

  
                    tableViewBody.innerHTML = "<tr><td colspan='30' class='text-center text-muted'>No switch OTB records found</td></tr>";
                });
            
            if (btnView) {
                btnView.addEventListener('click', search);
            }
            if (btnExport) {
                btnExport.addEventListener('click', exportTXN);
            }
        }

        let search = function () {
            var segmentCode = segmentDropdown.value;
            var cate = categoryDropdown.value;
            var brandCode = brandDropdown.value;
            var vendorCode = vendorDropdown.value;
            var Switchtype = typeDropdown.value;
            let OTByear = yearDropdown.value;
            let OTBmonth = monthDropdown.value;
            let OTBcompany = companyDropdown.value;


            var formData = new FormData();
            formData.append('OTBtype', Switchtype);
            formData.append('OTByear', OTByear);
            formData.append('OTBmonth', OTBmonth);
            formData.append('OTBCompany', OTBcompany);
            formData.append('OTBCategory', cate);
            formData.append('OTBSegment', segmentCode);
            formData.append('OTBBrand', brandCode);
            formData.append('OTBVendor', vendorCode);


            $.ajax({
                url: 'Handler/DataOTBHandler.ashx?action=obtswitchlistbyfilter',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    if (tableViewBody) tableViewBody.innerHTML = response; // <-- Added check
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }

        // *** ADDED: exportTXN function ***
        let exportTXN = function () {
            console.log("Export TXN (Switching) clicked");
            // Build query string from filters
            var params = new URLSearchParams();
            params.append('action', 'exportswitchingtxn'); // This action needs to be implemented in DataOTBHandler.ashx.vb
            params.append('OTBtype', typeDropdown.value);
            params.append('OTByear', yearDropdown.value);
            params.append('OTBmonth', monthDropdown.value);
            params.append('OTBCompany', companyDropdown.value);
            params.append('OTBCategory', categoryDropdown.value);
            params.append('OTBSegment', segmentDropdown.value);
            params.append('OTBBrand', brandDropdown.value);
            params.append('OTBVendor', vendorDropdown.value);

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
        
        }

        let InitSegment = function (segmentDropdown) {
            $.ajax({
                url: 'Handler/MasterDataHandler.ashx?action=SegmentMSList',
                type: 'POST',
                processData: false,
                contentType: false,
                success: function (response) {
                    if (segmentDropdown) segmentDropdown.innerHTML = response; // <-- Added check
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
                    if (yearDropdown) yearDropdown.innerHTML = response; // <-- Added check
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
                    if (monthDropdown) monthDropdown.innerHTML = response; // <-- Added check
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
                    if (companyDropdown) companyDropdown.innerHTML = response; // <-- Added check
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
                    if (categoryDropdown) categoryDropdown.innerHTML = response; // <-- Added check
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
                    if (brandDropdown) brandDropdown.innerHTML = response; // <-- Added check
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
                    if (vendorDropdown) vendorDropdown.innerHTML = response; // <-- Added check
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
                    if (vendorDropdown) vendorDropdown.innerHTML = response; // <-- Added check
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                }
            });
        }

        document.addEventListener('DOMContentLoaded', initial);

    </script>

</body>
</html>