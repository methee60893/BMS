<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="master_brand.aspx.vb" Inherits="BMS.master_brand" %>

<!DOCTYPE html>
<html lang="en">
<head runat="server">
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BMS - Master Brand</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
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
            background: #FF99CC;
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
                <h3><i class="bi bi-building"></i> BMS</h3>
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
                        <li><a href="master_vendor.aspx" class="menu-link">Master Vendor</a></li>
                        <li><a href="master_brand.aspx" class="menu-link active">Master Brand</a></li>
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
                    <h1 class="page-title" id="pageTitle">Master Brand</h1>
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
                            <label class="form-label">Brand Code</label>
                            <input type="text" id="txtSearchCode" class="form-control" placeholder="Enter brand code" autocomplete="off" />
                        </div>
                        <div class="col-md-7">
                            <label class="form-label">Brand Name</label>
                            <input type="text" id="txtSearchName" class="form-control" placeholder="Enter brand name" autocomplete="off" />
                        </div>
                    </div>

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

                <div class="export-section">
                    <asp:Button ID="btnExport" runat="server" Text="📊 Export to Excel" CssClass="btn btn-export btn-custom" OnClick="btnExport_Click" />
                </div>

                <div class="table-container">
                     <table id="brandTable" class="table table-hover mb-0">
                        <thead class="bg-light text-dark">
                            <tr>
                                 <th style="width: 200px;">Brand Code</th>
                                <th style="width: 400px;">Brand Name</th>
                                <th style="width: 100px;">Status</th>
                                <th style="width: 180px;">Actions</th>
                             </tr>
                        </thead>
                        <tbody id="brandTableBody">
                            <tr>
                                 <td colspan="3" class="text-center text-muted">Loading...</td>
                            </tr>
                        </tbody>
                     </table>
                </div>

                <div class="modal fade" id="brandModal" tabindex="-1" aria-labelledby="brandModalLabel" data-bs-backdrop="static" data-bs-keyboard="false">
                    <div class="modal-dialog modal-lg">
                        <div class="modal-content">
                             <div class="modal-header">
                                <h5 class="modal-title" id="brandModalLabel">Create New Brand</h5>
                                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                             </div>
                            <div class="modal-body">
                                <input type="hidden" id="hdnEditMode" value="create" />
                                <input type="hidden" id="hdnOriginalBrandCode" value="" />

                                <div class="row g-3">
                                    <div class="col-md-4">
                                         <label class="form-label">Brand Code <span class="required">*</span></label>
                                        <input type="text" id="txtModalCode" class="form-control" placeholder="Enter brand code" maxlength="50" autocomplete="off" />
                                     </div>
                                    <div class="col-md-8">
                                        <label class="form-label">Brand Name <span class="required">*</span></label>
                                         <input type="text" id="txtModalName" class="form-control" placeholder="Enter brand name" maxlength="255" autocomplete="off" />
                                    </div>
                                    <div class="col-12">
                                        <div class="form-check form-switch mt-2">
                                            <input class="form-check-input" type="checkbox" role="switch" id="chkModalActive" checked>
                                            <label class="form-check-label" for="chkModalActive">Active</label>
                                        </div>
                                    </div>
                                 </div>
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
        <script type="text/javascript">

            let brandModal; // ตัวแปรสำหรับ Bootstrap Modal Instance

            function showLoading(show) {
                console.log(show ? "Loading..." : "Done.");
                // (สามารถเพิ่ม Logic การแสดง Loading overlay ที่นี่)
            }

            function clearModalForm() {
                $('#hdnEditMode').val('create');
                $('#hdnOriginalBrandCode').val('');
                $('#brandModalLabel').text('Create New Brand');
                $('#txtModalCode').val('').prop('readonly', false); // เปิดให้แก้ไข Code
                $('#txtModalName').val('');
                $('#chkModalActive').prop('checked', true);
            }

            function loadBrandTable() {
                showLoading(true);
                const searchCode = $('#txtSearchCode').val();
                const searchName = $('#txtSearchName').val();

                $.ajax({
                    type: "POST",
                    url: "Handler/MasterDataHandler.ashx?action=getbrandlisthtml",
                    data: {
                        searchCode: searchCode,
                        searchName: searchName
                    },
                    success: function (html) {
                        $('#brandTableBody').html(html);
                        showLoading(false);
                    },
                    error: function (xhr) {
                        showLoading(false);
                        alert('Error loading brand data: ' + xhr.responseText);
                    }
                });
            }

            // --- DOM Ready (เมื่อหน้าเว็บโหลดเสร็จ) ---
            $(document).ready(function () {

                // 1. Initialize Modal
                brandModal = new bootstrap.Modal(document.getElementById('brandModal'));

                // 2. "Create" Button Click
                $('#btnShowCreateModal').on('click', function () {
                    clearModalForm();
                    brandModal.show();
                });

                // 3. "Edit" Button Click (Event Delegation)
                $('#brandTableBody').on('click', '.btn-edit-brand', function () {
                    const btn = $(this);
                    clearModalForm();

                    $('#hdnEditMode').val('edit');
                    $('#brandModalLabel').text('Edit Brand');

                    const code = btn.data('code');
                    $('#hdnOriginalBrandCode').val(code);

                    $('#txtModalCode').val(code).prop('readonly', true); // ปิดการแก้ไข Code
                    $('#txtModalName').val(btn.data('name'));

                    const isActive = btn.data('active') === 'true' || btn.data('active') === true;
                    $('#chkModalActive').prop('checked', isActive);

                    brandModal.show();
                });

                // 4. "Delete" Button Click (Event Delegation)
                $('#brandTableBody').on('click', '.btn-delete-brand', function () {
                    const btn = $(this);
                    const code = btn.data('code');
                    const name = btn.data('name');

                    if (!confirm(`Are you sure you want to delete Brand: ${code} (${name})?`)) {
                        return;
                    }

                    showLoading(true);
                    $.ajax({
                        type: "POST",
                        url: "Handler/MasterDataHandler.ashx?action=deletebrand",
                        data: { brandCode: code },
                        dataType: "json",
                        success: function (response) {
                            showLoading(false);
                            if (response.success) {
                                alert(response.message);
                                loadBrandTable(); // โหลดตารางใหม่
                            } else {
                                alert('Error: ' + response.message);
                            }
                        },
                        error: function (xhr) {
                            showLoading(false);
                            alert('Fatal error deleting brand: ' + xhr.responseText);
                        }
                    });
                });

                // 5. "Save" Button (in Modal) Click
                $('#btnModalSave').on('click', function () {
                    const mode = $('#hdnEditMode').val();
                    const brandData = {
                        editMode: mode,
                        code: $('#txtModalCode').val(),
                        originalCode: $('#hdnOriginalBrandCode').val(),
                        name: $('#txtModalName').val(),
                        isActive: $('#chkModalActive').is(':checked')
                    };

                    if (!brandData.code || !brandData.name) {
                        alert('Brand Code and Brand Name are required!');
                        return;
                    }

                    showLoading(true);
                    $.ajax({
                        type: "POST",
                        url: "Handler/MasterDataHandler.ashx?action=savebrand",
                        data: brandData,
                        dataType: "json",
                        success: function (response) {
                            showLoading(false);
                            if (response.success) {
                                alert(response.message);
                                brandModal.hide();
                                loadBrandTable(); // โหลดตารางใหม่
                            } else {
                                alert('Error: ' + response.message);
                            }
                        },
                        error: function (xhr) {
                            showLoading(false);
                            alert('Fatal error saving brand: ' + xhr.responseText);
                        }
                    });
                });

                // 6. "View" Button Click
                $('#btnViewTable').on('click', function () {
                    loadBrandTable();
                });

                // 7. "Clear Filter" Button Click
                $('#btnClearFilter').on('click', function () {
                    $('#txtSearchCode').val('');
                    $('#txtSearchName').val('');
                    loadBrandTable();
                });

                // 8. Initial Load
                loadBrandTable();
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

                if (document.readyState === 'loading') {
                    document.addEventListener('DOMContentLoaded', init);
                } else {
                    init();
                }
            })();
        </script>
    </form>
</body>
</html>