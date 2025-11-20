<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="otbRemaining.aspx.vb" Inherits="BMS.otbRemaining" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - OTB Remaining</title>
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
                    <li><a href="draftPO.aspx" class="menu-link">Draft PO</a></li>
                    <li><a href="matchActualPO.aspx" class="menu-link">Match Actual PO</a></li>
                    <li><a href="actualPO.aspx" class="menu-link">Actual PO</a></li>
                </ul>
            </li>
            <li class="menu-item">
                <a href="otbRemaining.aspx" class="menu-link active">
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
                <h1 class="page-title" id="pageTitle">KBMS - OTB Remaining</h1>
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
                OTB Remaining
            </div>

            <!-- Filter Box -->
            <div class="filter-box">
                <div class="filter-title">
                    OTB Remaining
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
                    <div class="col-md-3">
                        <label class="form-label">Company</label>
                        <select id="DDCompany" class="form-select">
                            </select>
                    </div>
                    <div class="col-md-3">
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
                        <button type="button" id="btnClearFilter" class="btn btn-clear btn-custom me-2">
                            <i class="bi bi-x-circle"></i> Clear filter
                        </button>
                        <button type="button" id="btnView" class="btn btn-view btn-custom me-2">
                            <i class="bi bi-eye"></i> View
                        </button>
                        <button type="button" id="btnTxnSum" class="btn btn-view btn-custom me-2">
                            <i class="bi bi-eye"></i> Export Summary
                        </button>
                        <button type="button" id="btnTxnMovement" class="btn btn-export btn-custom">
                            <i class="bi bi-file-earmark-excel"></i> Export TXN
                        </button>
                    </div>
                </div>
            </div>

            <!-- Detail Summary -->
            <div class="detail-box">
                    <div class="detail-title">Detail</div>
                    
                    <%-- *** MODIFIED: Added IDs to all detail-value divs *** --%>
                    <div class="detail-row">
                        <div class="detail-label">Budget Approved (Original)</div>
                        <div class="detail-value" id="detail_BudgetApproved_Original">
                            0.00 THB
                            <%-- <a href="#" class="detail-link ms-3">Click history</a> --%>
                        </div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Revised Diff</div>
                        <div class="detail-value" id="detail_RevisedDiff">0.00 THB</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Extra</div>
                        <div class="detail-value" id="detail_Budget_Extra">
                            0.00 THB
                            <%-- <a href="#" class="detail-link ms-3">Click history</a> --%>
                        </div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Switch in</div>
                        <div class="detail-value" id="detail_Budget_SwitchIn">
                            0.00 THB
                            <%-- <a href="#" class="detail-link ms-3">Click history</a> --%>
                        </div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Balance in</div>
                        <div class="detail-value" id="detail_Budget_BalanceIn">0.00 THB</div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Carry in</div>
                        <div class="detail-value" id="detail_Budget_CarryIn">0.00 THB</div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Switch in</div>
                        <div class="detail-value" id="detail_Budget_SwitchOut">
                            0.00 THB
                            <%-- <a href="#" class="detail-link ms-3">Click history</a> --%>
                        </div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Balance out</div>
                        <div class="detail-value" id="detail_Budget_BalanceOut">0.00 THB</div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Carry out</div>
                        <div class="detail-value" id="detail_Budget_CarryOut">0.00 THB</div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label"><strong>Total Budget Approved</strong></div>
                        <div class="detail-value"><strong id="detail_TotalBudgetApproved">0.00 THB</strong></div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label">Total Actual/Draft</div>
                        <div class="detail-value" id="detail_TotalPO_Usage">
                            0.00 THB
                            <%-- <a href="#" class="detail-link ms-3">Click history</a> --%>
                        </div>
                    </div>

                    <div class="detail-row">
                        <div class="detail-label"><strong>Remaining</strong></div>
                        <div class="detail-value"><strong id="detail_Remaining">0.00 THB</strong></div>
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
            
            document.querySelectorAll('.menu-link').forEach(link => {
                link.classList.remove('active');
            });
            
            event.currentTarget.classList.add('active');
            document.getElementById('pageTitle').textContent = pageName;
            
            if (window.innerWidth <= 768) {
                toggleSidebar();
            }
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

        // ===================================================
        // ===== NEW SCRIPT BLOCK FOR MASTER DATA =========
        // ===================================================

        // Cache elements
        let yearDropdown = document.getElementById("DDYear");
        let monthDropdown = document.getElementById("DDMonth");
        let companyDropdown = document.getElementById("DDCompany");
        let segmentDropdown = document.getElementById("DDSegment");
        let categoryDropdown = document.getElementById("DDCategory");
        let brandDropdown = document.getElementById("DDBrand");
        let vendorDropdown = document.getElementById("DDVendor");
        let btnClearFilter = document.getElementById("btnClearFilter");
        let btnView = document.getElementById("btnView");
        
        

        // Initialize
        let initial = function () {

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
            segmentDropdown.addEventListener('change', changeVendor);
            btnClearFilter.addEventListener('click', function () {
                // Clear filter fields
                $("DDYear").val(null).trigger('change');
                $("DDMonth").val(null).trigger('change');
                $("DDCompany").val(null).trigger('change');
                $("DDSegment").val(null).trigger('change');
                $("DDCategory").val(null).trigger('change');
                $("DDBrand").val(null).trigger('change');
                $("DDVendor").val(null).trigger('change');
                // Re-initialize vendor dropdown (to show all)
                InitVendor(vendorDropdown);
                // Clear table
               
            });
            btnView.addEventListener('click', search);
            $('#btnTxnMovement').on('click', exportOTBMovement);
            $('#btnTxnSum').on('click', exportCategorySum);
        }

        // Main function to load all master data
        let InitMSData = function () {
            InitSegment(segmentDropdown);
            InitCategoty(categoryDropdown);
            InitBrand(brandDropdown);
            InitVendor(vendorDropdown);
            InitMSYear(yearDropdown);
            InitMonth(monthDropdown);
            InitCompany(companyDropdown);
        }

        // Search function (placeholder - needs handler logic)
        let search = function () {
            var segmentCode = segmentDropdown.value;
            var cate = categoryDropdown.value;
            var brandCode = brandDropdown.value;
            var vendorCode = vendorDropdown.value;
            let OTByear = yearDropdown.value;
            let OTBmonth = monthDropdown.value;
            let OTBcompany = companyDropdown.value;

            var formData = new FormData();
            formData.append('OTByear', OTByear);
            formData.append('OTBmonth', OTBmonth);
            formData.append('OTBCompany', OTBcompany);
            formData.append('OTBCategory', cate);
            formData.append('OTBSegment', segmentCode);
            formData.append('OTBBrand', brandCode);
            formData.append('OTBVendor', vendorCode);

            // Check if all fields are selected
            if (!OTByear) {
                alert("Please select year to view the report.");
                return;
            }

            console.log("Searching with filters...", Object.fromEntries(formData));
            // Show loading state
          
            // Clear detail box
            populateDetailBox(null); // Clear detail box

            $.ajax({
                url: 'Handler/DataOTBRemainingHandler.ashx', // <-- Call the NEW handler
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                dataType: 'json', // Expect JSON back
                success: function (response) {
                    // response will be a DataSet: { "detail": [..], "otherRemaining": [..] }
                    if (response.detail && response.detail.length > 0) {
                        populateDetailBox(response.detail[0]);
                    } else {
                        // No exact match found for detail, but still show "Other"
                        populateDetailBox(null);
                    }

                 
                },
                error: function (xhr, status, error) {
                    console.log('Error getlist data: ' + error, xhr.responseText);
                    let errorMsg = "Error loading data.";
                    try {
                        // Try to parse error from our handler
                        let errResponse = JSON.parse(xhr.responseText);
                        if (errResponse.message) errorMsg = errResponse.message;
                    } catch (e) { }

                 
                    populateDetailBox(null); // Clear detail on error
                }
            });
        }

        // ===== NEW: Helper function to format numbers =====
        function formatTHB(number) {
            if (number === null || number === undefined) {
                return "0.00 THB";
            }
            // Convert to number before formatting
            return parseFloat(number).toLocaleString('en-US', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            }) + " THB";
        }

        // ===== NEW: Helper function to populate detail box =====
        function populateDetailBox(data) {
            if (data) {
                // Populate all fields from the 'detail' (Result Set 1)
                document.getElementById('detail_BudgetApproved_Original').textContent = formatTHB(data.BudgetApproved_Original);
                document.getElementById('detail_RevisedDiff').textContent = formatTHB(data.RevisedDiff); // <-- ADDED
                document.getElementById('detail_Budget_Extra').textContent = formatTHB(data.Budget_Extra);
                document.getElementById('detail_Budget_SwitchIn').textContent = formatTHB(data.Budget_SwitchIn);
                document.getElementById('detail_Budget_BalanceIn').textContent = formatTHB(data.Budget_BalanceIn);
                document.getElementById('detail_Budget_CarryIn').textContent = formatTHB(data.Budget_CarryIn);
                document.getElementById('detail_Budget_SwitchOut').textContent = formatTHB(data.Budget_SwitchOut);
                document.getElementById('detail_Budget_BalanceOut').textContent = formatTHB(data.Budget_BalanceOut);
                document.getElementById('detail_Budget_CarryOut').textContent = formatTHB(data.Budget_CarryOut);
                document.getElementById('detail_TotalBudgetApproved').textContent = formatTHB(data.TotalBudgetApproved);

                // --- Logic from image: TotalPO_Usage = TotalDraftPO + TotalActualPO ---
                // (SP ส่งมา 3 ค่า เราจะใช้ TotalPO_Usage)
                document.getElementById('detail_TotalPO_Usage').textContent = formatTHB(data.TotalPO_Usage);

                document.getElementById('detail_Remaining').textContent = formatTHB(data.Remaining);
            } else {
                // Clear all values to 0.00 THB if no data
                document.getElementById('detail_BudgetApproved_Original').textContent = "0.00 THB";
                document.getElementById('detail_RevisedDiff').textContent = "0.00 THB"; // <-- ADDED
                document.getElementById('detail_Budget_Extra').textContent = "0.00 THB";
                document.getElementById('detail_Budget_SwitchIn').textContent = "0.00 THB";
                document.getElementById('detail_Budget_BalanceIn').textContent = "0.00 THB";
                document.getElementById('detail_Budget_CarryIn').textContent = "0.00 THB";
                document.getElementById('detail_Budget_SwitchOut').textContent = "0.00 THB";
                document.getElementById('detail_Budget_BalanceOut').textContent = "0.00 THB";
                document.getElementById('detail_Budget_CarryOut').textContent = "0.00 THB";
                document.getElementById('detail_TotalBudgetApproved').textContent = "0.00 THB";
                document.getElementById('detail_TotalPO_Usage').textContent = "0.00 THB";
                document.getElementById('detail_Remaining').textContent = "0.00 THB";
            }
        }

        // ===== NEW: Helper function to populate 'Other Remaining' table =====
      

        // --- Master Data Loaders (using '...MSList' actions) ---

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

        // Dependent Dropdown
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

        let exportOTBMovement = function () {
            console.log("Export OTB Movement clicked");

            // รับค่า Filter
            let OTByear = yearDropdown.value;
            let OTBmonth = monthDropdown.value;
            let OTBcompany = companyDropdown.value;
            let OTBCategory = categoryDropdown.value;
            let OTBSegment = segmentDropdown.value;
            let OTBBrand = brandDropdown.value;
            let OTBVendor = vendorDropdown.value;

            if (!OTByear) {
                alert("Please select a Year to export the report.");
                return;
            }

            var params = new URLSearchParams();
            params.append('action', 'exportOTBMovement'); // <--- Action ใหม่
            params.append('OTByear', OTByear);

            if (OTBmonth) params.append('OTBmonth', OTBmonth);
            if (OTBcompany) params.append('OTBCompany', OTBcompany);
            if (OTBCategory) params.append('OTBCategory', OTBCategory);
            if (OTBSegment) params.append('OTBSegment', OTBSegment);
            if (OTBBrand) params.append('OTBBrand', OTBBrand);
            if (OTBVendor) params.append('OTBVendor', OTBVendor);

            // Trigger Download
            window.location.href = 'Handler/DataOTBHandler.ashx?' + params.toString();
        }

        let exportCategorySum = function () {
            console.log("Export Category Summary clicked");

            // รับค่า Filter
            let OTByear = yearDropdown.value;
            let OTBmonth = monthDropdown.value;
            let OTBcompany = companyDropdown.value;
            let OTBCategory = categoryDropdown.value;
            let OTBSegment = segmentDropdown.value;
            // Brand และ Vendor ไม่จำเป็นต้องส่งไปกรอง เพราะ Report นี้ดูภาพรวม แต่ถ้าส่งไปก็ไม่เสียหาย

            if (!OTByear) {
                alert("Please select a Year to export the report.");
                return;
            }

            var params = new URLSearchParams();
            params.append('action', 'exportsummarycategory'); // <--- Action ใหม่
            params.append('OTByear', OTByear);

            if (OTBmonth) params.append('OTBmonth', OTBmonth);
            if (OTBcompany) params.append('OTBCompany', OTBcompany);
            if (OTBCategory) params.append('OTBCategory', OTBCategory);
            if (OTBSegment) params.append('OTBSegment', OTBSegment);

            // Trigger Download
            window.location.href = 'Handler/DataOTBHandler.ashx?' + params.toString();
        }

        // Initialize on page load
        document.addEventListener('DOMContentLoaded', initial);
    </script>
</body>
</html>