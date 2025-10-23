<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="dashboard.aspx.vb" Inherits="BMS.dashboard" %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BMS - Dashboard</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <style>
        :root {
            --primary-blue: #0B56A4;
            --sidebar-bg: #2c3e50;
            --sidebar-hover: #34495e;
            --orange-accent: #ff6b35;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #ecf0f1;
            overflow-x: hidden;
        }

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
        }

        .sidebar-menu {
            list-style: none;
            padding: 15px 0;
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
        }

        .main-wrapper {
            width: 100%;
            min-height: 100vh;
        }

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
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .page-title {
            color: #2c3e50;
            margin: 0;
            font-size: 1.6rem;
            font-weight: 600;
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

        .content-area {
            padding: 30px;
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }

        .stat-card {
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            transition: transform 0.3s;
        }

        .stat-card:hover {
            transform: translateY(-5px);
        }

        .stat-icon {
            width: 60px;
            height: 60px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.8rem;
            margin-bottom: 15px;
        }

        .stat-card.blue .stat-icon {
            background: rgba(11, 86, 164, 0.1);
            color: var(--primary-blue);
        }

        .stat-card.green .stat-icon {
            background: rgba(40, 167, 69, 0.1);
            color: #28a745;
        }

        .stat-card.orange .stat-icon {
            background: rgba(255, 107, 53, 0.1);
            color: var(--orange-accent);
        }

        .stat-card.purple .stat-icon {
            background: rgba(156, 39, 176, 0.1);
            color: #9c27b0;
        }

        .stat-label {
            color: #6c757d;
            font-size: 0.9rem;
            margin-bottom: 5px;
        }

        .stat-value {
            font-size: 2rem;
            font-weight: 700;
            color: #2c3e50;
        }

        .chart-container {
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }

        .chart-title {
            font-size: 1.2rem;
            font-weight: 600;
            color: #2c3e50;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #dee2e6;
        }

        .recent-activity {
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }

        .activity-item {
            display: flex;
            align-items: start;
            padding: 15px 0;
            border-bottom: 1px solid #e9ecef;
        }

        .activity-item:last-child {
            border-bottom: none;
        }

        .activity-icon {
            width: 40px;
            height: 40px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 15px;
            font-size: 1.1rem;
        }

        .activity-icon.blue {
            background: rgba(11, 86, 164, 0.1);
            color: var(--primary-blue);
        }

        .activity-icon.green {
            background: rgba(40, 167, 69, 0.1);
            color: #28a745;
        }

        .activity-icon.orange {
            background: rgba(255, 107, 53, 0.1);
            color: var(--orange-accent);
        }

        .activity-content {
            flex: 1;
        }

        .activity-title {
            font-weight: 600;
            color: #2c3e50;
            margin-bottom: 5px;
        }

        .activity-time {
            color: #6c757d;
            font-size: 0.85rem;
        }

        @media (max-width: 768px) {
            .content-area {
                padding: 15px;
            }

            .stats-grid {
                grid-template-columns: 1fr;
            }
        }

        .sidebar::-webkit-scrollbar {
            width: 6px;
        }

        .sidebar::-webkit-scrollbar-thumb {
            background: #34495e;
            border-radius: 3px;
        }
    </style>
</head>
<body>
    <div class="sidebar-overlay" id="sidebarOverlay" onclick="toggleSidebar()"></div>

    <div class="sidebar" id="sidebar">
        <div class="sidebar-header">
            <h3><i class="bi bi-building"></i> BMS</h3>
            <button class="close-sidebar" onclick="toggleSidebar()">
                <i class="bi bi-x-lg"></i>
            </button>
        </div>
        <ul class="sidebar-menu">
            <li class="menu-item">
                <a href="#" class="menu-link" onclick="toggleSubmenu(event, 'otbPlan')">
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

    <div class="main-wrapper">
        <div class="top-navbar">
            <div class="d-flex align-items-center gap-3">
                <button class="menu-toggle" onclick="toggleSidebar()">
                    <i class="bi bi-list"></i>
                </button>
                <h1 class="page-title">BMS</h1>
            </div>
            <div class="d-flex align-items-center gap-3">
                <span class="d-none d-md-inline">Welcome, Admin</span>
                <div class="user-avatar">
                    <i class="bi bi-person-circle"></i>
                </div>
            </div>
        </div>

        <div class="content-area">
            <div class="stats-grid">
                <div class="stat-card blue">
                    <div class="stat-icon">
                        <i class="bi bi-file-earmark-text"></i>
                    </div>
                    <div class="stat-label">Total OTB Plans</div>
                    <div class="stat-value">124</div>
                </div>
                <div class="stat-card green">
                    <div class="stat-icon">
                        <i class="bi bi-check-circle"></i>
                    </div>
                    <div class="stat-label">Approved</div>
                    <div class="stat-value">98</div>
                </div>
                <div class="stat-card orange">
                    <div class="stat-icon">
                        <i class="bi bi-clock-history"></i>
                    </div>
                    <div class="stat-label">Pending</div>
                    <div class="stat-value">26</div>
                </div>
                <div class="stat-card purple">
                    <div class="stat-icon">
                        <i class="bi bi-currency-dollar"></i>
                    </div>
                    <div class="stat-label">Total Budget (THB)</div>
                    <div class="stat-value">12.5M</div>
                </div>
            </div>

            <div class="row">
                <div class="col-lg-8 mb-4">
                    <div class="chart-container">
                        <div class="chart-title">Monthly Budget Overview</div>
                        <canvas id="budgetChart" height="80"></canvas>
                    </div>
                </div>
                <div class="col-lg-4 mb-4">
                    <div class="recent-activity">
                        <div class="chart-title">Recent Activities</div>
                        <div class="activity-item">
                            <div class="activity-icon blue">
                                <i class="bi bi-file-plus"></i>
                            </div>
                            <div class="activity-content">
                                <div class="activity-title">New Draft OTB Created</div>
                                <div class="activity-time">2 hours ago</div>
                            </div>
                        </div>
                        <div class="activity-item">
                            <div class="activity-icon green">
                                <i class="bi bi-check-circle"></i>
                            </div>
                            <div class="activity-content">
                                <div class="activity-title">OTB Plan Approved</div>
                                <div class="activity-time">5 hours ago</div>
                            </div>
                        </div>
                        <div class="activity-item">
                            <div class="activity-icon orange">
                                <i class="bi bi-arrow-left-right"></i>
                            </div>
                            <div class="activity-content">
                                <div class="activity-title">Switching Transaction Completed</div>
                                <div class="activity-time">1 day ago</div>
                            </div>
                        </div>
                        <div class="activity-item">
                            <div class="activity-icon blue">
                                <i class="bi bi-file-earmark-check"></i>
                            </div>
                            <div class="activity-content">
                                <div class="activity-title">PO Matched Successfully</div>
                                <div class="activity-time">2 days ago</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
        function toggleSidebar() {
            document.getElementById('sidebar').classList.toggle('active');
            document.getElementById('sidebarOverlay').classList.toggle('active');
        }

        function toggleSubmenu(event, submenuId) {
            event.preventDefault();
            event.stopPropagation();
            const submenu = document.getElementById(submenuId);
            const menuLink = event.currentTarget;
            submenu.classList.toggle('show');
            menuLink.classList.toggle('expanded');
        }

        // Chart
        const ctx = document.getElementById('budgetChart').getContext('2d');
        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
                datasets: [{
                    label: 'Approved Budget',
                    data: [1.2, 1.9, 1.5, 2.1, 1.8, 2.3],
                    backgroundColor: 'rgba(11, 86, 164, 0.8)',
                }, {
                    label: 'Remaining',
                    data: [0.5, 0.8, 0.6, 0.9, 0.7, 1.1],
                    backgroundColor: 'rgba(40, 167, 69, 0.8)',
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: {
                        position: 'bottom',
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                return value + 'M';
                            }
                        }
                    }
                }
            }
        });
    </script>
</body>
</html>