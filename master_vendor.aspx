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
</head>
<body>
    <form id="form1" runat="server">
        <!-- Sidebar Overlay -->
        <div class="sidebar-overlay" id="sidebarOverlay"></div>

        <!-- Sidebar -->
        <div class="sidebar" id="sidebar">
            <div class="sidebar-header">
                <h3><a class="text-decoration-none text-white" href="dashboard.aspx" ><i class="bi bi-building"></i> KBMS</a></h3>
                <div class="close-sidebar" id="closeSidebarBtn">
                    <i class="bi bi-x-lg"></i>
                </div>
            </div>
            <ul class="sidebar-menu">
                <li class="menu-item">
                    <a href="#" class="menu-link" data-submenu="otbPlan">
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
                <li class="menu-item"><a href="default.aspx" class="menu-link"><i class="bi bi-box-arrow-left"></i> Logout</a></li>
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
                    <span class="d-none d-md-inline">Welcome, <%= HttpUtility.JavaScriptStringEncode(Session("fullname").ToString()) %></span>
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
                                <input type="hidden" id="hdVendorId" value="" />
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
                        $('#hdVendorId').val(btn.data('vendorid'));
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
                    $('#hdVendorId').val(btn.data('vendorid'));
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
                    const venid = btn.data('vendorid');
                    const code = btn.data('code');
                    const name = btn.data('name');

                    if (!confirm(`Are you sure you want to delete Vendor: ${code} (${name})?`)) {
                        return;
                    }

                    showLoading(true);
                    $.ajax({
                        type: "POST",
                        url: "Handler/MasterDataHandler.ashx?action=deleteVendor",
                        data: { vendorCode: code, vendorId: venid },
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
                        vendorId: $('#hdVendorId').val(),
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