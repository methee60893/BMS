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
                <li class="menu-item"><a href="dashboard.aspx" class="menu-link"><i class="bi bi-house"></i> Dashboard</a></li>
                <li class="menu-item"><a href="manage_users.aspx" class="menu-link active"><i class="bi bi-people"></i> Manage Users</a></li>
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
    </script>
</body>
</html>