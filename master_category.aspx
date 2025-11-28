<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="master_category.aspx.vb" Inherits="BMS.master_category" %>

<!DOCTYPE html>
<html lang="en">
<head runat="server">
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - Master Category</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="style/theme.css" rel="stylesheet">
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
                        <li><a href="master_vendor.aspx" class="menu-link">Master Vendor</a></li>
                        <li><a href="master_brand.aspx" class="menu-link">Master Brand</a></li>
                        <li><a href="master_category.aspx" class="menu-link active">Master Category</a></li>
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
                    <h1 class="page-title" id="pageTitle">Master Category</h1>
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
                            <label class="form-label">Category Code</label>
                            <input type="text" id="txtSearchCode" class="form-control" placeholder="Enter category code" autocomplete="off" />
                        </div>
                        <div class="col-md-7">
                            <label class="form-label">Category Name</label>
                            <input type="text" id="txtSearchName" class="form-control" placeholder="Enter category name" autocomplete="off" />
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
                    <table id="categoryTable" class="table table-hover mb-0">
                        <thead class="bg-light text-dark">
                            <tr>
                                <th style="width: 200px;">Category Code</th>
                                <th style="width: 400px;">Category Name</th>
                                <!-- (START) ADDED COLUMN -->
                                <th style="width: 100px;">Status</th>
                                <!-- (END) ADDED COLUMN -->
                                <th style="width: 180px;">Actions</th>
                            </tr>
                        </thead>
                        <tbody id="categoryTableBody">
                            <tr>
                                <!-- (MODIFIED) Colspan increased -->
                                <td colspan="4" class="text-center text-muted">Loading...</td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                <div class="modal fade" id="categoryModal" tabindex="-1" aria-labelledby="categoryModalLabel" data-bs-backdrop="static" data-bs-keyboard="false">
                    <div class="modal-dialog modal-lg">
                        <div class="modal-content">
                            <div class="modal-header">
                                <h5 class="modal-title" id="categoryModalLabel">Create New Category</h5>
                                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                            </div>
                            <div class="modal-body">
                                <input type="hidden" id="hdnEditMode" value="create" />
                                <input type="hidden" id="hdnOriginalCategoryCode" value="" />

                                <div class="row g-3">
                                    <div class="col-md-4">
                                        <label class="form-label">Category Code <span class="required">*</span></label>
                                        <input type="text" id="txtModalCode" class="form-control" placeholder="Enter category code" maxlength="50" autocomplete="off" />
                                    </div>
                                    <div class="col-md-8">
                                        <label class="form-label">Category Name <span class="required">*</span></label>
                                        <input type="text" id="txtModalName" class="form-control" placeholder="Enter category name" maxlength="255" autocomplete="off" />
                                    </div>
                                    <!-- (START) ADDED TOGGLE SWITCH -->
                                    <div class="col-12">
                                        <div class="form-check form-switch mt-2">
                                            <input class="form-check-input" type="checkbox" role="switch" id="chkModalActiveCategory" checked>
                                            <label class="form-check-label" for="chkModalActiveCategory">Active</label>
                                        </div>
                                    </div>
                                    <!-- (END) ADDED TOGGLE SWITCH -->
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

            let categoryModal; // ตัวแปรสำหรับ Bootstrap Modal Instance

            function showLoading(show) {
                console.log(show ? "Loading..." : "Done.");
                // (สามารถเพิ่ม Logic การแสดง Loading overlay ที่นี่)
            }
            function clearModalForm() {
                $('#hdnEditMode').val('create');
                $('#hdnOriginalCategoryCode').val('');
                $('#categoryModalLabel').text('Create New Category');
                $('#txtModalCode').val('').prop('readonly', false); // เปิดให้แก้ไข Code
                $('#txtModalName').val('');
                $('#chkModalActiveCategory').prop('checked', true); // (MODIFIED) Reset toggle
            }

            function loadCategoryTable() {
                showLoading(true);
                const searchCode = $('#txtSearchCode').val();
                const searchName = $('#txtSearchName').val();

                $.ajax({
                    type: "POST",
                    url: "Handler/MasterDataHandler.ashx?action=getcategorylisthtml",
                    data: {
                        searchCode: searchCode,
                        searchName: searchName
                    },
                    success: function (html) {
                        $('#categoryTableBody').html(html);
                        showLoading(false);
                    },
                    error: function (xhr) {
                        showLoading(false);
                        alert('Error loading category data: ' + xhr.responseText);
                    }
                });
            }

            $(document).ready(function () {

                // 1. Initialize Modal
                categoryModal = new bootstrap.Modal(document.getElementById('categoryModal'));

                // 2. "Create" Button Click
                $('#btnShowCreateModal').on('click', function () {
                    clearModalForm();
                    categoryModal.show();
                });

                // 3. "Edit" Button Click (Event Delegation)
                $('#categoryTableBody').on('click', '.btn-edit-category', function () {
                    const btn = $(this);
                    clearModalForm();

                    $('#hdnEditMode').val('edit');
                    $('#categoryModalLabel').text('Edit Category');

                    const code = btn.data('code');
                    $('#hdnOriginalCategoryCode').val(code);

                    $('#txtModalCode').val(code).prop('readonly', true); // ปิดการแก้ไข Code
                    $('#txtModalName').val(btn.data('name'));

                    // (START) MODIFIED: Read isActive status
                    const isActive = btn.data('active') === 'true' || btn.data('active') === true;
                    $('#chkModalActiveCategory').prop('checked', isActive);
                    // (END) MODIFIED

                    categoryModal.show();
                });

                // 4. "Delete" Button Click (Event Delegation)
                $('#categoryTableBody').on('click', '.btn-delete-category', function () {
                    const btn = $(this);
                    const code = btn.data('code');
                    const name = btn.data('name');

                    if (!confirm(`Are you sure you want to delete Category: ${code} (${name})?`)) {
                        return;
                    }

                    showLoading(true);
                    $.ajax({
                        type: "POST",
                        url: "Handler/MasterDataHandler.ashx?action=deletecategory",
                        data: { categoryCode: code },
                        dataType: "json",
                        success: function (response) {
                            showLoading(false);
                            if (response.success) {
                                alert(response.message);
                                loadCategoryTable(); // โหลดตารางใหม่
                            } else {
                                alert('Error: ' + response.message);
                            }
                        },
                        error: function (xhr) {
                            showLoading(false);
                            alert('Fatal error deleting category: ' + xhr.responseText);
                        }
                    });
                });

                // 5. "Save" Button (in Modal) Click
                $('#btnModalSave').on('click', function () {
                    const mode = $('#hdnEditMode').val();
                    const categoryData = {
                        editMode: mode,
                        code: $('#txtModalCode').val(),
                        originalCode: $('#hdnOriginalCategoryCode').val(),
                        name: $('#txtModalName').val(),
                        // (START) MODIFIED: Send isActive status
                        isActive: $('#chkModalActiveCategory').is(':checked')
                        // (END) MODIFIED
                    };

                    if (!categoryData.code || !categoryData.name) {
                        alert('Category Code and Category Name are required!');
                        return;
                    }

                    showLoading(true);
                    $.ajax({
                        type: "POST",
                        url: "Handler/MasterDataHandler.ashx?action=savecategory",
                        data: categoryData,
                        dataType: "json",
                        success: function (response) {
                            showLoading(false);
                            if (response.success) {
                                alert(response.message);
                                categoryModal.hide();
                                loadCategoryTable(); // โหลดตารางใหม่
                            } else {
                                alert('Error: ' + response.message);
                            }
                        },
                        error: function (xhr) {
                            showLoading(false);
                            alert('Fatal error saving category: ' + xhr.responseText);
                        }
                    });
                });

                // 6. "View" Button Click
                $('#btnViewTable').on('click', function () {
                    loadCategoryTable();
                });

                // 7. "Clear Filter" Button Click
                $('#btnClearFilter').on('click', function () {
                    $('#txtSearchCode').val('');
                    $('#txtSearchName').val('');
                    loadCategoryTable();
                });

                // 8. Initial Load
                loadCategoryTable();
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
