<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="manage_users.aspx.vb" Inherits="BMS.manage_users" %>

<!DOCTYPE html>
<html lang="en">
<head runat="server">
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KBMS - Manage Users</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="style/theme.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
</head>
<body>
    <form id="form1" runat="server">
        <div class="sidebar-overlay" id="sidebarOverlay"></div>
        
        <div class="sidebar" id="sidebar">
            <div class="sidebar-header">
                <h3><a class="text-decoration-none text-white" href="dashboard.aspx"><i class="bi bi-building"></i> KBMS</a></h3>
                <div class="close-sidebar" id="closeSidebarBtn"><i class="bi bi-x-lg"></i></div>
            </div>
            <ul class="sidebar-menu">
                <li class="menu-item" id="grpmenuOTBPlan" runat="server">
                    <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'otbPlan')">
                        <i class="bi bi-clipboard-data"></i>
                        <span>OTB Plan / Revise</span>
                        <i class="bi bi-chevron-down"></i>
                    </a>
                    <ul class="submenu" id="otbPlan">
                        <li id="menuDraftOTBPlan" runat="server"><a href="draftOTB.aspx" class="menu-link">Draft OTB Plan</a></li>
                        <li id="menuApprovedOTBPlan" runat="server"><a href="approvedOTB.aspx" class="menu-link">Approved OTB Plan</a></li>
                    </ul>
                </li>
                <li class="menu-item" id="grpmenuOTBSwitching" runat="server">
                    <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'otbSwitching')">
                        <i class="bi bi-arrow-left-right"></i>
                        <span>OTB Switching</span>
                        <i class="bi bi-chevron-down"></i>
                    </a>
                    <ul class="submenu" id="otbSwitching">
                        <li id="menuCreateOTBSwitching" runat="server"><a href="createOTBswitching.aspx" class="menu-link">Create OTB Switching</a></li>
                        <li id="menuSwitchingTransaction" runat="server"><a href="transactionOTBSwitching.aspx" class="menu-link">Switching Transaction</a></li>
                    </ul>
                </li>
                <li class="menu-item" id="grpmenuPO" runat="server">
                    <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'po')">
                        <i class="bi bi-file-earmark-text"></i>
                        <span>PO</span>
                        <i class="bi bi-chevron-down"></i>
                    </a>
                    <ul class="submenu" id="po">
                        <li id="menuCreateDraftPO" runat="server"><a href="createDraftPO.aspx" class="menu-link">Create Draft PO</a></li>
                        <li id="menuDraftPO" runat="server"><a href="draftPO.aspx" class="menu-link">Draft PO</a></li>
                        <li id="menuMatchActualPO" runat="server"><a href="matchActualPO.aspx" class="menu-link">Match Actual PO</a></li>
                        <li id="menuActualPO" runat="server"><a href="actualPO.aspx" class="menu-link">Actual PO</a></li>
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
                        <li id="menuVendor" runat="server"><a href="master_vendor.aspx" class="menu-link">Master Vendor</a></li>
                        <li id="menuBrand" runat="server"><a href="master_brand.aspx" class="menu-link">Master Brand</a></li>
                        <li id="menuCategory" runat="server"><a href="master_category.aspx" class="menu-link">Master Category</a></li>
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
                        <li id="menuManageUsers" runat="server"><a href="manage_users.aspx" class="menu-link active">Manage Users</a></li>
                    </ul>
                </li>
                <li class="menu-item"><a href="default.aspx" class="menu-link"><i class="bi bi-box-arrow-left"></i> Logout</a></li>
            </ul>
        </div>

        <div class="main-wrapper">
            <div class="top-navbar">
                <div class="d-flex align-items-center gap-3">
                    <div class="menu-toggle" id="menuToggleBtn"><i class="bi bi-list"></i></div>
                    <h1 class="page-title">Manage Users & Roles</h1>
                </div>
                <div class="user-info">
                    <span class="d-none d-md-inline">Admin Panel</span>
                    <div class="user-avatar"><i class="bi bi-person-circle"></i></div>
                </div>
            </div>

            <div class="content-area">
                <div class="master-box">
                    <div class="master-title">User List</div>
                    <div class="row g-3 mb-4">
                        <div class="col-md-6">
                            <div class="input-group">
                                <span class="input-group-text"><i class="bi bi-search"></i></span>
                                <input type="text" id="txtSearch" class="form-control" placeholder="Search by Username or Name...">
                                <button type="button" id="btnSearch" class="btn btn-primary">Search</button>
                            </div>
                        </div>
                    </div>

                    <div class="table-responsive">
                        <table class="table table-hover align-middle">
                            <thead class="table-light">
                                <tr>
                                    <th>ID</th>
                                    <th>Username</th>
                                    <th>Full Name</th>
                                    <th>Email</th>
                                    <th>Current Roles</th>
                                    <th>Status</th>
                                    <th class="text-center">Action</th>
                                </tr>
                            </thead>
                            <tbody id="tblUsersBody">
                                </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <div class="modal fade" id="userModal" tabindex="-1">
            <div class="modal-dialog">
                <div class="modal-content">
                    <div class="modal-header bg-primary text-white">
                        <h5 class="modal-title"><i class="bi bi-person-gear"></i> Edit User Role</h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">
                        <input type="hidden" id="hdnUserId">
                        <div class="mb-3">
                            <label class="form-label">Username</label>
                            <input type="text" id="txtUsername" class="form-control" readonly disabled>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Full Name</label>
                            <input type="text" id="txtFullName" class="form-control" readonly disabled>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Assign Role</label>
                            <select id="ddlRole" class="form-select">
                                </select>
                        </div>
                        <div class="form-check form-switch">
                            <input class="form-check-input" type="checkbox" id="chkActive">
                            <label class="form-check-label" for="chkActive">Active User</label>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                        <button type="button" id="btnSaveUser" class="btn btn-primary">Save Changes</button>
                    </div>
                </div>
            </div>
        </div>

    </form>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>

    <script>
        var userModal;

        $(document).ready(function () {
            userModal = new bootstrap.Modal(document.getElementById('userModal'));

            // Load initial data
            loadUsers();
            loadRoles();

            // Search button
            $('#btnSearch').click(function () {
                loadUsers($('#txtSearch').val());
            });

            // Save button
            $('#btnSaveUser').click(function () {
                saveUser();
            });
        });

        function loadUsers(search = '') {
            $.ajax({
                url: 'Handler/UserHandler.ashx?action=getUsers',
                type: 'GET',
                data: { search: search },
                dataType: 'json',
                success: function (data) {
                    var html = '';
                    if (data.length === 0) {
                        html = '<tr><td colspan="7" class="text-center">No users found</td></tr>';
                    } else {
                        $.each(data, function (i, item) {
                            var statusBadge = item.IsActive ? '<span class="badge bg-success">Active</span>' : '<span class="badge bg-secondary">Inactive</span>';
                            html += `<tr>
                                        <td>${item.UserID}</td>
                                        <td>${item.Username}</td>
                                        <td>${item.FullName}</td>
                                        <td>${item.Email}</td>
                                        <td><span class="badge bg-info text-dark">${item.Roles || '-'}</span></td>
                                        <td>${statusBadge}</td>
                                        <td class="text-center">
                                            <button type="button" class="btn btn-sm btn-warning" onclick="editUser(${item.UserID}, '${item.Username}', '${item.FullName}', ${item.IsActive})">
                                                <i class="bi bi-pencil"></i> Edit
                                            </button>
                                        </td>
                                     </tr>`;
                        });
                    }
                    $('#tblUsersBody').html(html);
                }
            });
        }

        function loadRoles() {
            $.ajax({
                url: 'Handler/UserHandler.ashx?action=getRoles',
                type: 'GET',
                dataType: 'json',
                success: function (data) {
                    var options = '<option value="">-- Select Role --</option>';
                    $.each(data, function (i, item) {
                        options += `<option value="${item.id}">${item.text}</option>`;
                    });
                    $('#ddlRole').html(options);
                }
            });
        }

        function editUser(id, username, fullname, isActive) {
            $('#hdnUserId').val(id);
            $('#txtUsername').val(username);
            $('#txtFullName').val(fullname);
            $('#chkActive').prop('checked', isActive);
            
            // Note: For simplicity, this demo assumes single role assignment. 
            // In a real multi-role scenario, you'd need to fetch the user's specific roles first.
            $('#ddlRole').val(''); 
            
            userModal.show();
        }

        function saveUser() {
            var userId = $('#hdnUserId').val();
            var roleId = $('#ddlRole').val();
            var isActive = $('#chkActive').is(':checked');

            if (!roleId) {
                alert('Please select a role');
                return;
            }

            $.ajax({
                url: 'Handler/UserHandler.ashx?action=saveUser',
                type: 'POST',
                data: {
                    userId: userId,
                    roleId: roleId,
                    isActive: isActive
                },
                success: function (response) {
                    alert('User updated successfully');
                    userModal.hide();
                    loadUsers($('#txtSearch').val());
                },
                error: function () {
                    alert('Error saving user');
                }
            });
        }
        
        // Sidebar Toggle Logic (Same as master pages)
        var sidebar = document.getElementById('sidebar');
        var overlay = document.getElementById('sidebarOverlay');
        var menuToggle = document.getElementById('menuToggleBtn');
        
        if(menuToggle){
            menuToggle.addEventListener('click', function(){
                sidebar.classList.toggle('active');
                overlay.classList.toggle('active');
            });
        }
        if(overlay){
            overlay.addEventListener('click', function(){
                sidebar.classList.remove('active');
                overlay.classList.remove('active');
            });
        }

        function toggleSubmenu(event, submenuId) {
            event.preventDefault();
            event.stopPropagation();

            var submenu = document.getElementById(submenuId);
            if (submenu) {
                submenu.classList.toggle('show');
            }

            if (event.currentTarget) {
                event.currentTarget.classList.toggle('expanded');
            }
        }
    </script>
</body>
</html>
