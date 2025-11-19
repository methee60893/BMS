<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="master_vendor.aspx.vb" Inherits="BMS.master_vendor" %>

<!DOCTYPE html>
<html lang="en">
<head runat="server">
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - Master Vendor</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="style/theme.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <style>
        :root {
            --primary-blue: #0B56A4;
            --sidebar-bg: #2c3e50;
            --sidebar-hover: #34495e;
            --orange-header: #D2691E;
            --light-blue-bg: #E6F2FF;
            --green-btn: #28a745;
            --yellow-btn: #FFC107;
            --red-btn: #dc3545;
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
            user-select: none;
        }

            .close-sidebar:hover {
                background: rgba(255,255,255,0.1);
            }

            .close-sidebar:active {
                background: rgba(255,255,255,0.2);
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
                background: var(--bms-active-menu);
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
            user-select: none;
        }

            .menu-toggle:hover {
                background: #094580;
                transform: scale(1.05);
            }

            .menu-toggle:active {
                transform: scale(0.95);
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

        /* Master Box */
        .master-box {
            background: white;
            border: 2px solid #dee2e6;
            border-radius: 8px;
            padding: 25px;
            margin-bottom: 25px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }

        .master-title {
            font-size: 1.3rem;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 3px solid var(--primary-blue);
        }

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

        /* Buttons */
        .btn-custom {
            padding: 10px 25px;
            border-radius: 6px;
            font-weight: 600;
            transition: all 0.3s ease;
            border: none;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            font-size: 0.95rem;
        }

        .btn-create {
            background: var(--primary-blue);
            color: white;
        }

            .btn-create:hover {
                background: #094580;
                transform: translateY(-2px);
            }

        .btn-clear {
            background: #17a2b8;
            color: white;
        }

            .btn-clear:hover {
                background: #138496;
                transform: translateY(-2px);
            }

        .btn-view {
            background: var(--primary-blue);
            color: white;
        }

            .btn-view:hover {
                background: #094580;
                transform: translateY(-2px);
            }

        .btn-submit {
            background: var(--green-btn);
            color: white;
        }

            .btn-submit:hover {
                background: #218838;
                transform: translateY(-2px);
            }

        .btn-export {
            background: var(--primary-blue);
            color: white;
        }

            .btn-export:hover {
                background: #094580;
                transform: translateY(-2px);
            }

        .btn-edit {
            background: var(--yellow-btn);
            color: #000;
            padding: 6px 15px;
            font-size: 0.85rem;
        }

            .btn-edit:hover {
                background: #e0a800;
            }

        .btn-delete {
            background: var(--red-btn);
            color: white;
            padding: 6px 15px;
            font-size: 0.85rem;
        }

            .btn-delete:hover {
                background: #c82333;
            }

        /* Create Section */
        .create-section {
            background: var(--primary-blue);
            color: white;
            padding: 12px 20px;
            font-weight: 700;
            font-size: 1.1rem;
            border-radius: 6px 6px 0 0;
            margin: -25px -25px 20px -25px;
        }

        /* Export Section */
        .export-section {
            display: flex;
            justify-content: flex-end;
            margin-bottom: 15px;
        }

        /* Data Table */
        .table-container {
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            overflow-x: auto;
        }

        .table {
            margin: 0;
            font-size: 0.9rem;
            white-space: nowrap;
        }

            .table thead {
                background: #f8f9fa;
                color: #2c3e50;
            }

                .table thead th {
                    padding: 14px 12px;
                    font-weight: 700;
                    border-bottom: 2px solid #dee2e6;
                    vertical-align: middle;
                }

            .table tbody td {
                padding: 12px;
                vertical-align: middle;
                border-bottom: 1px solid #e9ecef;
            }

            .table tbody tr:hover {
                background: #f8f9fa;
            }

        .required {
            color: red;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <!-- Sidebar Overlay -->
        <div class="sidebar-overlay" id="sidebarOverlay"></div>

        <!-- Sidebar -->
        <div class="sidebar" id="sidebar">
            <div class="sidebar-header">
                <h3><i class="bi bi-building"></i>KBMS</h3>
                <div class="close-sidebar" id="closeSidebarBtn">
                    <i class="bi bi-x-lg"></i>
                </div>
            </div>
            <ul class="sidebar-menu">
                <li class="menu-item">
                    <a href="#" class="menu-link" data-submenu="otbPlan">
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
                    <a href="#" class="menu-link" data-submenu="otbSwitching">
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
                    <a href="#" class="menu-link" data-submenu="po">
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
                    <a href="#" class="menu-link" data-submenu="master">
                        <i class="bi bi-database"></i>
                        <span>Master File</span>
                        <i class="bi bi-chevron-down"></i>
                    </a>
                    <ul class="submenu" id="master">
                        <li><a href="master_vendor.aspx" class="menu-link active">Master Vendor</a></li>
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
                    <div class="menu-toggle" id="menuToggleBtn">
                        <i class="bi bi-list"></i>
                    </div>
                    <h1 class="page-title" id="pageTitle">Master Vendor</h1>
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
                <!-- Search/Filter Box -->
                <div class="master-box">
                    <div class="master-title">Search & Filter</div>

                    <div class="row g-3 mb-4">
                        <div class="col-md-3">
                            <label class="form-label">Vendor Code</label>
                            <input id="txtSearchCode" class="form-control" placeholder="Enter vendor code" autocomplete="off" />
                        </div>
                        <div class="col-md-4">
                            <label class="form-label">Vendor Name</label>
                            <input id="txtSearchName" class="form-control" placeholder="Enter vendor name" autocomplete="off" />
                        </div>
                        <div class="col-md-3">
                            <label class="form-label">Segment</label>
                            <select id="ddlSearchSegment" class="form-select">
                                <option value="">-- All Segments --</option>
                            </select>
                        </div>
                        <div class="col-md-2">
                        </div>
                    </div>

                    <!-- Action Buttons -->
                    <div class="row">
                        <div class="col-12 text-end">
                            <button type="button" id="btnShowCreateModal" class="btn btn-create btn-custom me-2">
                                <i class="bi bi-plus-circle"></i>Create
                           
                            </button>
                            <button type="button" id="btnClearFilter" class="btn btn-clear btn-custom me-2">
                                <i class="bi bi-x-circle"></i>Clear filter
                           
                            </button>
                            <button type="button" id="btnViewTable" class="btn btn-view btn-custom">
                                <i class="bi bi-eye"></i>View
                           
                            </button>
                        </div>
                    </div>
                </div>

                <!-- Export Button -->
                <div class="export-section">
                    <asp:Button ID="btnExport" runat="server" Text="📊 Export to Excel" CssClass="btn btn-export btn-custom" OnClick="btnExport_Click" />
                </div>
                <!-- Data Table -->
                <div class="table-container">
                    <table id="vendorTable" class="table table-hover mb-0">
                        <thead class="bg-light text-dark">
                            <tr>
                                <th style="width: 120px;">Vendor Code</th>
                                <th style="width: 200px;">Vendor Name</th>
                                <th style="width: 80px;">CCY</th>
                                <th style="width: 120px;">Payment Term Code</th>
                                <th style="width: 150px;">Payment Term</th>
                                <th style="width: 120px;">Segment Code</th>
                                <th style="width: 150px;">Segment</th>
                                <th style="width: 100px;">Incoterm</th>
                                <!-- (START) ADDED COLUMN -->
                                <th style="width: 100px;">Status</th>
                                <!-- (END) ADDED COLUMN -->
                                <th style="width: 180px;">Actions</th>
                            </tr>
                        </thead>
                        <tbody id="vendorTableBody">
                            <tr>
                                <!-- (MODIFIED) Colspan increased -->
                                <td colspan="10" class="text-center text-muted">Loading...</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <div class="modal fade" id="vendorModal" tabindex="-1" aria-labelledby="vendorModalLabel" data-bs-backdrop="static" data-bs-keyboard="false">
                    <div class="modal-dialog modal-lg">
                        <div class="modal-content">
                            <div class="modal-header">
                                <h5 class="modal-title" id="vendorModalLabel">Create New Vendor</h5>
                                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                            </div>
                            <div class="modal-body">
                                <input type="hidden" id="hdnEditMode" value="create" />
                                <input type="hidden" id="hdnOriginalVendorCode" value="" />

                                <div class="row g-3">
                                    <div class="col-md-4">
                                        <label class="form-label">Vendor Code <span class="required">*</span></label>
                                        <input type="text" id="txtModalCode" class="form-control" placeholder="Enter vendor code" maxlength="50" autocomplete="off" />
                                    </div>
                                    <div class="col-md-8">
                                        <label class="form-label">Vendor Name <span class="required">*</span></label>
                                        <input type="text" id="txtModalName" class="form-control" placeholder="Enter vendor name" maxlength="255" autocomplete="off" />
                                    </div>
                                </div>
                                <div class="row g-3">
                                    <div class="col-md-3">
                                        <label class="form-label">CCY</label>
                                        <select id="ddlModalCCY" class="form-select" data-placeholder="-- Select CCY --">
                                        </select>
                                    </div>
                                    <div class="col-md-3">
                                        <label class="form-label">Payment Term Code</label>
                                        <input type="text" id="txtModalPaymentTermCode" class="form-control" placeholder="Enter code" maxlength="50" autocomplete="off" />
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">Payment Term</label>
                                        <input type="text" id="txtModalPaymentTerm" class="form-control" placeholder="Enter payment term" maxlength="255" autocomplete="off" />
                                    </div>
                                </div>
                                <div class="row g-3">
                                    <div class="col-md-9">
                                        <label class="form-label">Segment</label>
                                        <select id="ddlModalSegment" class="form-select" data-placeholder="-- Select Segment --">
                                            </select>
                                    </div>
                                    <div class="col-md-3">
                                        <label class="form-label">Incoterm</label>
                                        <input type="text" id="txtModalIncoterm" class="form-control" placeholder="e.g., FOB, CIF" maxlength="50" autocomplete="off" />
                                    </div>
                                </div>
                                <!-- (START) ADDED TOGGLE SWITCH -->
                                <div class="row g-3 mt-2">
                                     <div class="col-12">
                                        <div class="form-check form-switch">
                                            <input class="form-check-input" type="checkbox" role="switch" id="chkModalActiveVendor" checked>
                                            <label class="form-check-label" for="chkModalActiveVendor">Active</label>
                                        </div>
                                    </div>
                                </div>
                                <!-- (END) ADDED TOGGLE SWITCH -->
                            </div>
                            <div class="modal-footer">
                                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                                <button type="button" id="btnModalSave" class="btn btn-primary">
                                    <i class="bi bi-check-circle"></i>Save Changes
               
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
        <script type="text/javascript">

            let vendorModal; // ตัวแปรสำหรับ Bootstrap Modal Instance

            function showLoading(show) {
                // (คุณสามารถเพิ่ม/ใช้ฟังก์ชัน Loading ที่สวยงามได้ที่นี่)
                console.log(show ? "Loading..." : "Done.");
            }

            function clearModalForm() {
                $('#hdnEditMode').val('create');
                $('#hdnOriginalVendorCode').val('');
                $('#vendorModalLabel').text('Create New Vendor');
                $('#txtModalCode').val('').prop('readonly', false); // เปิดให้แก้ไข Code
                $('#txtModalName').val('');
                $('#txtModalIncoterm').val('');
                $('#txtModalPaymentTermCode').val('');
                $('#txtModalPaymentTerm').val('');

                $('#ddlModalCCY').val(null).trigger('change');
                $('#ddlModalSegment').val(null).trigger('change');
                $('#chkModalActiveVendor').prop('checked', true); // (MODIFIED) Reset toggle
            }

            function loadVendorTable() {
                showLoading(true);
                const searchCode = $('#txtSearchCode').val();
                const searchName = $('#txtSearchName').val();
                const searchSegment = $('#ddlSearchSegment').val();
                console.log('searchCode : ' + searchCode + ', searchName : ' + searchName + ', searchSegment : ' + searchSegment);
                $.ajax({
                    type: "POST",
                    url: "Handler/MasterDataHandler.ashx?action=getVendorListHtml",
                    data: {
                        searchCode: searchCode,
                        searchName: searchName,
                        segmentCode: searchSegment
                    },
                    success: function (html) {
                        $('#vendorTableBody').html(html);
                        showLoading(false);
                    },
                    error: function (xhr) {
                        showLoading(false);
                        alert('Error loading vendor data: ' + xhr.responseText);
                    }
                });
            }

            // (MODIFIED: แก้ไขให้ฟังก์ชัน return $.ajax promise)
            function InitMSCCY(element) {
                return $.ajax({ // <--- **[CHANGES 1]**
                    url: 'Handler/MasterDataHandler.ashx?action=CCYMSList',
                    type: 'POST',
                    data: { addAll: false },
                    success: function (response) {
                        element.innerHTML = response;
                        $(element).prepend('<option value=""></option>').val(null);
                    },
                    error: function (xhr) { console.error('Error loading CCY:', xhr.responseText); }
                });
            }

            // (MODIFIED: แก้ไขให้ฟังก์ชัน return $.ajax promise)
            function InitMSSegment(element) {
                return $.ajax({ // <--- **[CHANGES 2]**
                    url: 'Handler/MasterDataHandler.ashx?action=SegmentMSList',
                    type: 'POST',
                    data: { addAll: false },
                    success: function (response) {
                        element.innerHTML = response;
                        $(element).prepend('<option value=""></option>').val(null);
                    },
                    error: function (xhr) { console.error('Error loading Segment:', xhr.responseText); }
                });
            }
            function InitMSSegmentSearch(element) {
                return $.ajax({ // <--- **[CHANGES 2]**
                    url: 'Handler/MasterDataHandler.ashx?action=SegmentMSList',
                    type: 'POST',
                    data: { addAll: false },
                    success: function (response) {
                        element.innerHTML = response;
                    },
                    error: function (xhr) { console.error('Error loading Segment:', xhr.responseText); }
                });
            }

            // --- DOM Ready (เมื่อหน้าเว็บโหลดเสร็จ) ---
            $(document).ready(function () {

                $('#ddlModalCCY').select2({
                    theme: "bootstrap-5",
                    dropdownParent: $("#vendorModal"), // (สำคัญมากสำหรับ Modal)
                    allowClear: true
                });
                $('#ddlModalSegment').select2({
                    theme: "bootstrap-5",
                    dropdownParent: $("#vendorModal"),
                    allowClear: true
                });

                // 3. (MODIFIED: เรียกใช้ฟังก์ชันและเก็บค่า Promise)
                var ccyPromise = InitMSCCY(document.getElementById('ddlModalCCY'));
                var segmentPromise = InitMSSegment(document.getElementById('ddlModalSegment'));

                // 4. (MODIFIED: ใช้ $.when() เพื่อรอให้ AJAX ทั้ง 2 ตัวทำงานเสร็จก่อน)
                //    (โค้ดส่วนนี้จะทำงาน "หลังจาก" ที่ Dropdown มีข้อมูลแล้วเท่านั้น)
                $.when(ccyPromise, segmentPromise).done(function () {

                    console.log("Master data for modal (CCY, Segment) loaded successfully.");

                    // 5. (MOVED: ย้าย Event "Edit" มาไว้ข้างใน .done())
                    //    (ตอนนี้ Event จะถูกผูกเมื่อ Dropdown พร้อมใช้งานแล้ว)
                    $('#vendorTableBody').on('click', '.btn-edit-vendor', function () {
                        const btn = $(this);
                        clearModalForm();

                        $('#hdnEditMode').val('edit');
                        $('#vendorModalLabel').text('Edit Vendor');

                        const code = btn.data('code');
                        $('#hdnOriginalVendorCode').val(code);
                        $('#txtModalCode').val(code).prop('readonly', true);
                        $('#txtModalName').val(btn.data('name'));

                        $('#txtModalPaymentTermCode').val(btn.data('term-code'));
                        $('#txtModalPaymentTerm').val(btn.data('term'));
                        $('#txtModalIncoterm').val(btn.data('incoterm'));

                        // (โค้ดส่วนนี้จะทำงานได้อย่างถูกต้องแล้ว)
                        $('#ddlModalCCY').val(btn.data('ccy')).trigger('change');
                        $('#ddlModalSegment').val(btn.data('seg-code')).trigger('change');
                        
                        // (START) MODIFIED: Read isActive status
                        const isActive = btn.data('active') === 'true' || btn.data('active') === true;
                        $('#chkModalActiveVendor').prop('checked', isActive);
                        // (END) MODIFIED

                        vendorModal.show();
                    });

                }).fail(function () {
                    alert("Critical error: Failed to load master data (CCY/Segment) for the edit modal. Please refresh the page.");
                });


                // 6. (Event ที่ไม่เกี่ยวข้องกับ Modal Master Data สามารถอยู่ที่เดิมได้)
                $('#btnShowCreateModal').on('click', function () {
                    clearModalForm();
                    // (Reset Select2 ไปที่ค่าว่าง)
                    $('#ddlModalCCY').val(null).trigger('change');
                    $('#ddlModalSegment').val(null).trigger('change');
                    vendorModal.show();
                });

                // 1. Initialize Modal
                vendorModal = new bootstrap.Modal(document.getElementById('vendorModal'));

                // 2. "Create" Button Click
                // (ใช้ ID ที่เราตั้งใหม่ใน <button type="button">)
                $('#btnShowCreateModal').on('click', function () {
                    clearModalForm();
                    vendorModal.show();
                });

                // 3. "Edit" Button Click (Event Delegation)
                // (เราดักฟังที่ GridView แล้วหาปุ่ม .btn-edit-vendor ที่ถูกคลิก)
                $('#vendorTableBody').on('click', '.btn-edit-vendor', function () {
                    const btn = $(this);
                    clearModalForm();

                    $('#hdnEditMode').val('edit');
                    $('#vendorModalLabel').text('Edit Vendor');

                    const code = btn.data('code');
                    $('#hdnOriginalVendorCode').val(code);

                    $('#txtModalCode').val(code).prop('readonly', true);
                    $('#txtModalName').val(btn.data('name'));
                    $('#txtModalPaymentTermCode').val(btn.data('term-code'));
                    $('#txtModalPaymentTerm').val(btn.data('term'));
                    $('#txtModalIncoterm').val(btn.data('incoterm'));

                    $('#ddlModalCCY').val(btn.data('ccy'));
                    $('#ddlModalSegment').val(btn.data('seg-code'));
                    
                    // (START) MODIFIED: Read isActive status
                    const isActive = btn.data('active') === 'true' || btn.data('active') === true;
                    $('#chkModalActiveVendor').prop('checked', isActive);
                    // (END) MODIFIED

                    vendorModal.show();
                });

                // 4. "Delete" Button Click (Event Delegation)
                // *** อัปเดตตรงนี้: เปลี่ยน gvVendor.ClientID เป็น #vendorTableBody ***
                $('#vendorTableBody').on('click', '.btn-delete-vendor', function () {
                    const btn = $(this);
                    const code = btn.data('code');
                    const name = btn.data('name');

                    if (!confirm(`Are you sure you want to delete Vendor: ${code} (${name})?`)) {
                        return;
                    }

                    showLoading(true);
                    $.ajax({
                        type: "POST",
                        url: "Handler/MasterDataHandler.ashx?action=deleteVendor",
                        data: { vendorCode: code },
                        dataType: "json",
                        success: function (response) {
                            showLoading(false);
                            if (response.success) {
                                alert(response.message);
                                loadVendorTable();
                            } else {
                                alert('Error: ' + response.message);
                            }
                        },
                        error: function (xhr) {
                            showLoading(false);
                            alert('Fatal error deleting vendor: ' + xhr.responseText);
                        }
                    });
                });

                // 5. "Save" Button (in Modal) Click
                $('#btnModalSave').on('click', function () {
                    // (ส่วนนี้เหมือนเดิม)
                    const mode = $('#hdnEditMode').val();
                    const vendorData = {
                        editMode: mode,
                        code: $('#txtModalCode').val(),
                        originalCode: $('#hdnOriginalVendorCode').val(),
                        name: $('#txtModalName').val(),
                        paymentTermCode: $('#txtModalPaymentTermCode').val(),
                        paymentTerm: $('#txtModalPaymentTerm').val(),
                        segmentCode: $('#ddlModalSegment').val(),
                        segment: $('#ddlModalSegment option:selected').text().split(' - ')[1], // (MODIFIED) Get text better
                        ccy: $('#ddlModalCCY').val(),
                        incoterm: $('#txtModalIncoterm').val(),
                        // (START) MODIFIED: Send isActive status
                        isActive: $('#chkModalActiveVendor').is(':checked')
                        // (END) MODIFIED
                    };

                    if (!vendorData.code || !vendorData.name) {
                        alert('Vendor Code and Vendor Name are required!');
                        return;
                    }

                    showLoading(true);

                    $.ajax({
                        type: "POST",
                        url: "Handler/MasterDataHandler.ashx?action=saveVendor",
                        data: vendorData,
                        dataType: "json",
                        success: function (response) {
                            showLoading(false);
                            if (response.success) {
                                alert(response.message);
                                vendorModal.hide();
                                loadVendorTable();
                            } else {
                                alert('Error: ' + response.message);
                            }
                        },
                        error: function (xhr) {
                            showLoading(false);
                            alert('Fatal error saving vendor: ' + xhr.responseText);
                        }
                    });
                });

                // 6. "View" Button Click
                $('#btnViewTable').on('click', function () {
                    loadVendorTable();
                });

                // 7. "Clear Filter" Button Click
                $('#btnClearFilter').on('click', function () {
                    $('#txtSearchCode').val('');
                    $('#txtSearchName').val('');
                    $('#ddlSearchSegment').val('');
                    loadVendorTable();
                });

                // 8. Initial Load (โหลดข้อมูลตารางครั้งแรกเมื่อหน้าเว็บพร้อม)
                loadVendorTable();

            }); // <-- End of $(document).ready()

            // Wait for DOM to be ready
            (function () {
                // Toggle Sidebar Function
                function toggleSidebar() {
                    var sidebar = document.getElementById('sidebar');
                    var overlay = document.getElementById('sidebarOverlay');

                    if (sidebar && overlay) {
                        sidebar.classList.toggle('active');
                        overlay.classList.toggle('active');
                    }
                }

                // Toggle Submenu Function
                function toggleSubmenu(element, submenuId) {
                    var submenu = document.getElementById(submenuId);

                    if (submenu && element) {
                        submenu.classList.toggle('show');
                        element.classList.toggle('expanded');
                    }
                }

                // Initialize when DOM is ready
                if (document.readyState === 'loading') {
                    document.addEventListener('DOMContentLoaded', init);
                } else {
                    init();
                }

                function init() {
                    // Menu Toggle Button
                    var menuToggleBtn = document.getElementById('menuToggleBtn');
                    if (menuToggleBtn) {
                        menuToggleBtn.addEventListener('click', function (e) {
                            e.preventDefault();
                            e.stopPropagation();
                            toggleSidebar();
                        });
                    }

                    var ddlSearchSegment = InitMSSegmentSearch(document.getElementById('ddlSearchSegment'));

                    // Close Sidebar Button
                    var closeSidebarBtn = document.getElementById('closeSidebarBtn');
                    if (closeSidebarBtn) {
                        closeSidebarBtn.addEventListener('click', function (e) {
                            e.preventDefault();
                            e.stopPropagation();
                            toggleSidebar();
                        });
                    }

                    // Sidebar Overlay
                    var overlay = document.getElementById('sidebarOverlay');
                    if (overlay) {
                        overlay.addEventListener('click', function (e) {
                            e.preventDefault();
                            toggleSidebar();
                        });
                    }

                    // Submenu Links (those with data-submenu)
                    var submenuTriggers = document.querySelectorAll('.menu-link[data-submenu]');
                    submenuTriggers.forEach(function (link) {
                        link.addEventListener('click', function (e) {
                            e.preventDefault();
                            e.stopPropagation();

                            var submenuId = this.getAttribute('data-submenu');
                            var submenu = document.getElementById(submenuId);

                            if (submenu) {
                                submenu.classList.toggle('show');
                                this.classList.toggle('expanded');
                            }
                        });
                    });

                    // Close sidebar when clicking outside
                    document.addEventListener('click', function (e) {
                        var sidebar = document.getElementById('sidebar');
                        var menuToggle = document.getElementById('menuToggleBtn');

                        if (sidebar && menuToggle) {
                            var isClickInsideSidebar = sidebar.contains(e.target);
                            var isClickOnToggle = menuToggle.contains(e.target);

                            if (!isClickInsideSidebar && !isClickOnToggle) {
                                if (sidebar.classList.contains('active')) {
                                    toggleSidebar();
                                }
                            }
                        }
                    });
                }
            })();
        </script>
    </form>
</body>
</html>