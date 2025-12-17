<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="actualPO.aspx.vb" Inherits="BMS.actualPO" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - Actual PO</title>
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
                    <li id="menuCreateOTBSwitching" runat="server" ><a href="createOTBswitching.aspx" class="menu-link">Create OTB Switching</a></li>
                    <li id="menuSwitchingTransaction" runat="server" ><a href="transactionOTBSwitching.aspx" class="menu-link">Switching Transaction</a></li>
                </ul>
            </li>
            <li class="menu-item">
                <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'po')">
                    <i class="bi bi-file-earmark-text"></i>
                    <span>PO</span>
                    <i class="bi bi-chevron-down"></i>
                </a>
                <ul class="submenu" id="po">
                    <li id="menu" runat="server" ><a href="createDraftPO.aspx" class="menu-link">Create Draft PO</a></li>
                    <li id="menu" runat="server" ><a href="draftPO.aspx" class="menu-link">Draft PO</a></li>
                    <li id="menu" runat="server" ><a href="matchActualPO.aspx" class="menu-link">Match Actual PO</a></li>
                    <li id="menu" runat="server" ><a href="actualPO.aspx" class="menu-link active">Actual PO</a></li>
                </ul>
            </li>
            <li id="menu" runat="server" class="menu-item">
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
                <h1 class="page-title" id="pageTitle">KBMS - Actual PO</h1>
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
                Actual PO
            </div>
            <!-- Filter Box -->
            <div class="filter-box">
                <div class="filter-title">
                   Search Actual PO
                </div>

                <%-- *** MODIFIED: Added IDs to all controls *** --%>
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
                            <i class="bi bi-x-circle"></i> Clear filter
                        </button>
                        <button type="button" id="btnView" class="btn btn-view btn-custom">
                            <i class="bi bi-eye"></i> View
                        </button>
                    </div>
                </div>
            </div>

             <!-- Export Button -->
            <div class="export-section">
                <button type="button" id="btnExport" class="btn btn-export btn-custom">
                    <i class="bi bi-file-earmark-excel"></i> Export TXN
                </button>
            </div>

            <!-- Data Table -->
            <!-- Data Table -->
            <div class="table-container">
                <div class="table-responsive">
                    <table class="table table-hover mb-0">
                        <thead>
                            <tr>
                                <th>Actual PO Date</th>
                                <th>Actual PO no.</th>
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
                                <th>Draft PO Ref</th>
                                <th>Remark</th>
                                <th>Status</th>
                                <th>Status date</th>
                            </tr>
                        </thead>
                        <%-- *** MODIFIED: Added ID and placeholder *** --%>
                        <tbody id="actualPOTableBody">
                            <tr>
                                <td colspan="22" class="text-center text-muted p-4">
                                    Please use the filters and click "View" to see data.
                                </td>
                            </tr>
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

        // ==========================================
        // ===== NEW: actualPO.aspx SCRIPT LOGIC =====
        // ==========================================

        // --- Cache Filter Elements ---
        let ddYearFilter = document.getElementById('ddYearFilter');
        let ddMonthFilter = document.getElementById('ddMonthFilter');
        let ddCompanyFilter = document.getElementById('ddCompanyFilter');
        let ddCategoryFilter = document.getElementById('ddCategoryFilter');
        let ddSegmentFilter = document.getElementById('ddSegmentFilter');
        let ddBrandFilter = document.getElementById('ddBrandFilter');
        let ddVendorFilter = document.getElementById('ddVendorFilter');
        let btnView = document.getElementById('btnView');
        let btnClearFilter = document.getElementById('btnClearFilter');
        let btnExport = document.getElementById('btnExport');
        let actualPOTableBody = document.getElementById('actualPOTableBody');

        // --- Initializer ---
        function initial() {

            if (ddYearFilter) {
                $(ddYearFilter).select2({
                    theme: "bootstrap-5"
                });
            }
            if (ddMonthFilter) {
                $(ddMonthFilter).select2({
                    theme: "bootstrap-5"
                });
            }
            if (ddCompanyFilter) {
                $(ddCompanyFilter).select2({
                    theme: "bootstrap-5"
                });
            }
            if (ddCategoryFilter) {
                $(ddCategoryFilter).select2({
                    theme: "bootstrap-5"
                });
            }
            if (ddSegmentFilter) {
                $(ddSegmentFilter).select2({
                    theme: "bootstrap-5"
                });
            }
            if (ddBrandFilter) {
                $(ddBrandFilter).select2({
                    theme: "bootstrap-5"
                });
            }
            if (ddVendorFilter) {
                $(ddVendorFilter).select2({
                    theme: "bootstrap-5"
                });
            }

            InitMSData();
            
            // Add Event Listeners
            btnView.addEventListener('click', search);
            btnClearFilter.addEventListener('click', clearFilters);
            btnExport.addEventListener('click', exportTXN);
        }

        // --- Clear Filters ---
        function clearFilters() {
            $('#ddYearFilter').val(null).trigger('change');
            $('#ddMonthFilter').val(null).trigger('change');
            $('#ddCompanyFilter').val(null).trigger('change');
            $('#ddCategoryFilter').val(null).trigger('change');
            $('#ddSegmentFilter').val(null).trigger('change');
            $('#ddBrandFilter').val(null).trigger('change');
            $('#ddVendorFilter').val(null).trigger('change');
            InitVendor(ddVendorFilter); // Reset vendor list
            actualPOTableBody.innerHTML = "<tr><td colspan='22' class='text-center text-muted p-4'>Please use the filters and click 'View' to see data.</td></tr>";
        }

        // --- Search Function (AJAX Call) ---
        function search() {
            var formData = new FormData();
            formData.append('year', ddYearFilter.value);
            formData.append('month', ddMonthFilter.value);
            formData.append('company', ddCompanyFilter.value);
            formData.append('category', ddCategoryFilter.value);
            formData.append('segment', ddSegmentFilter.value);
            formData.append('brand', ddBrandFilter.value);
            formData.append('vendor', ddVendorFilter.value);

            // Show loading state
            actualPOTableBody.innerHTML = "<tr><td colspan='22' class='text-center text-muted p-4'><div class='spinner-border spinner-border-sm' role='status'></div> Loading data...</td></tr>";

            $.ajax({
                url: 'Handler/DataPOHandler.ashx?action=getActualPOList', // Call new action
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    actualPOTableBody.innerHTML = response; // Inject HTML
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error);
                    actualPOTableBody.innerHTML = `<tr><td colspan='22' class='text-center text-danger p-4'>Error loading data: ${xhr.responseText}</td></tr>`;
                }
            });
        }

        // --- Export Function ---
        function exportTXN() {
            console.log("Export TXN clicked");
            // Build query string from filters
            var params = new URLSearchParams();
            params.append('action', 'exportActualPO'); // Call new export action
            params.append('year', ddYearFilter.value);
            params.append('month', ddMonthFilter.value);
            params.append('company', ddCompanyFilter.value);
            params.append('category', ddCategoryFilter.value);
            params.append('segment', ddSegmentFilter.value);
            params.append('brand', ddBrandFilter.value);
            params.append('vendor', ddVendorFilter.value);

            // Use window.location to trigger file download (GET request)
            window.location.href = 'Handler/DataPOHandler.ashx?' + params.toString();
        }

        // ==========================================
        // --- Master Data Loaders ---
        // (Copied from other pages for consistency)
        // ==========================================

        function InitMSData() {
            InitSegment(ddSegmentFilter);
            InitCategoty(ddCategoryFilter);
            InitBrand(ddBrandFilter);
            InitVendor(ddVendorFilter);
            InitMSYear(ddYearFilter);
            InitMonth(ddMonthFilter);
            InitCompany(ddCompanyFilter);
            
        }

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

        // Run initializer on load
        document.addEventListener('DOMContentLoaded', initial);

    </script>
</body>
</html>