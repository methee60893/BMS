<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="default.aspx.vb" Inherits="BMS._default" %>

<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>BMS - Login</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
  <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
  <link rel="icon" href="data:,">
  
  <style>
    :root {
      --primary-color: #1a237e;
      --secondary-color: #3949ab;
      --accent-color: #00bcd4;
      --gold-accent: #ffc107;
      --bg-primary: #0f1419;
      --bg-secondary: #1e2328;
      --bg-card: #252932;
      --text-primary: #ffffff;
      --text-secondary: #b8bcc8;
      --border-color: #3a404a;
      --shadow-primary: 0 20px 40px rgba(0, 0, 0, 0.3);
      --shadow-card: 0 8px 24px rgba(0, 0, 0, 0.2);
    }

    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }

    body {
      background: linear-gradient(135deg, #0f1419 0%, #1a1d29 50%, #0f1419 100%);
      min-height: 100vh;
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      color: var(--text-primary);
      overflow: hidden;
      display: flex;
      align-items: center;
      justify-content: center;
      position: relative;
    }

    /* Animated background particles */
    .bg-particles {
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      overflow: hidden;
      z-index: 0;
      opacity: 0.3;
    }

    .particle {
      position: absolute;
      width: 2px;
      height: 2px;
      background: var(--accent-color);
      border-radius: 50%;
      animation: float linear infinite;
    }

    @keyframes float {
      0% {
        transform: translateY(100vh) translateX(0);
        opacity: 0;
      }
      10% {
        opacity: 1;
      }
      90% {
        opacity: 1;
      }
      100% {
        transform: translateY(-100vh) translateX(100px);
        opacity: 0;
      }
    }

    .login-container {
      position: relative;
      z-index: 1;
      width: 100%;
      max-width: 440px;
      padding: 20px;
    }

    .login-card {
      background: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 20px;
      padding: 3rem;
      box-shadow: var(--shadow-primary);
      position: relative;
      overflow: hidden;
      backdrop-filter: blur(10px);
    }

    .login-card::before {
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 3px;
      background: linear-gradient(90deg, var(--primary-color), var(--accent-color), var(--gold-accent));
    }

    .login-header {
      text-align: center;
      margin-bottom: 2.5rem;
    }

    .logo-container {
      width: 80px;
      height: 80px;
      margin: 0 auto 1.5rem;
      background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
      border-radius: 20px;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 8px 24px rgba(26, 35, 126, 0.4);
      position: relative;
    }

    .logo-container::after {
      content: '';
      position: absolute;
      inset: -2px;
      background: linear-gradient(45deg, var(--accent-color), var(--gold-accent));
      border-radius: 20px;
      z-index: -1;
      opacity: 0;
      transition: opacity 0.3s ease;
    }

    .login-card:hover .logo-container::after {
      opacity: 0.3;
    }

    .logo-container i {
      font-size: 2.5rem;
      color: var(--text-primary);
    }

    .login-title {
      font-size: 1.8rem;
      font-weight: 700;
      color: var(--text-primary);
      margin-bottom: 0.5rem;
      letter-spacing: -0.02em;
    }

    .login-subtitle {
      font-size: 0.95rem;
      color: var(--text-secondary);
      font-weight: 300;
      letter-spacing: 0.02em;
    }

    .form-group {
      margin-bottom: 1.5rem;
    }

    .form-label {
      color: var(--text-secondary);
      font-weight: 500;
      font-size: 0.85rem;
      margin-bottom: 0.6rem;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      display: block;
    }

    .input-wrapper {
      position: relative;
    }

    .input-icon {
      position: absolute;
      left: 1rem;
      top: 50%;
      transform: translateY(-50%);
      color: var(--text-secondary);
      font-size: 1.1rem;
      transition: color 0.3s ease;
      z-index: 2;
      pointer-events: none;
    }

    /* Override ASP.NET TextBox styles */
    .form-control,
    input[type="text"].form-control,
    input[type="password"].form-control {
      background: var(--bg-secondary) !important;
      border: 1px solid var(--border-color) !important;
      border-radius: 10px !important;
      color: var(--text-primary) !important;
      padding: 0.875rem 1rem 0.875rem 3rem !important;
      font-size: 0.95rem !important;
      transition: all 0.3s ease !important;
      width: 100% !important;
      height: auto !important;
    }

    .form-control:focus,
    input[type="text"].form-control:focus,
    input[type="password"].form-control:focus {
      background: var(--bg-secondary) !important;
      border-color: var(--accent-color) !important;
      box-shadow: 0 0 0 0.2rem rgba(0, 188, 212, 0.25) !important;
      color: var(--text-primary) !important;
      outline: none !important;
    }

    .form-control::placeholder {
      color: rgba(184, 188, 200, 0.5) !important;
    }

    .password-toggle {
      position: absolute;
      right: 1rem;
      top: 50%;
      transform: translateY(-50%);
      background: none;
      border: none;
      color: var(--text-secondary);
      cursor: pointer;
      padding: 0.5rem;
      transition: color 0.3s ease;
      z-index: 3;
    }

    .password-toggle:hover {
      color: var(--accent-color);
    }

    .remember-forgot {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 2rem;
    }

    .remember-me {
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }

    .remember-me input[type="checkbox"] {
      width: 18px;
      height: 18px;
      accent-color: var(--accent-color);
      cursor: pointer;
    }

    .remember-me label {
      color: var(--text-secondary);
      font-size: 0.9rem;
      cursor: pointer;
      margin: 0;
    }

    .forgot-link {
      color: var(--accent-color);
      text-decoration: none;
      font-size: 0.9rem;
      font-weight: 500;
      transition: color 0.3s ease;
    }

    .forgot-link:hover {
      color: var(--gold-accent);
    }

    /* Override ASP.NET Button styles */
    .btn-kp-primary,
    input[type="submit"].btn-kp-primary {
      background: linear-gradient(135deg, var(--primary-color), var(--secondary-color)) !important;
      border: none !important;
      border-radius: 10px !important;
      padding: 1rem 2rem !important;
      font-weight: 600 !important;
      color: var(--text-primary) !important;
      font-size: 1rem !important;
      transition: all 0.3s ease !important;
      box-shadow: var(--shadow-card) !important;
      position: relative !important;
      overflow: hidden !important;
      text-transform: uppercase !important;
      letter-spacing: 0.08em !important;
      width: 100% !important;
      cursor: pointer !important;
      height: auto !important;
    }

    .btn-kp-primary::before,
    input[type="submit"].btn-kp-primary::before {
      content: '';
      position: absolute;
      top: 0;
      left: -100%;
      width: 100%;
      height: 100%;
      background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.2), transparent);
      transition: left 0.5s ease;
    }

    .btn-kp-primary:hover::before,
    input[type="submit"].btn-kp-primary:hover::before {
      left: 100%;
    }

    .btn-kp-primary:hover,
    input[type="submit"].btn-kp-primary:hover {
      transform: translateY(-2px) !important;
      box-shadow: 0 12px 28px rgba(26, 35, 126, 0.5) !important;
      background: linear-gradient(135deg, var(--secondary-color), var(--primary-color)) !important;
    }

    .btn-kp-primary:active,
    input[type="submit"].btn-kp-primary:active {
      transform: translateY(0) !important;
    }

    .btn-kp-primary:disabled,
    input[type="submit"].btn-kp-primary:disabled {
      opacity: 0.6 !important;
      cursor: not-allowed !important;
      transform: none !important;
    }

    .divider {
      display: flex;
      align-items: center;
      margin: 2rem 0;
      color: var(--text-secondary);
      font-size: 0.85rem;
    }

    .divider::before,
    .divider::after {
      content: '';
      flex: 1;
      height: 1px;
      background: var(--border-color);
    }

    .divider span {
      padding: 0 1rem;
    }

    .alert-message {
      background: rgba(244, 67, 54, 0.1);
      border: 1px solid rgba(244, 67, 54, 0.3);
      color: #f44336;
      padding: 1rem;
      border-radius: 8px;
      font-size: 0.9rem;
      margin-bottom: 1.5rem;
      display: none;
      animation: slideDown 0.3s ease;
    }

    .alert-message.show {
      display: block;
    }

    @keyframes slideDown {
      from {
        opacity: 0;
        transform: translateY(-10px);
      }
      to {
        opacity: 1;
        transform: translateY(0);
      }
    }

    .footer-text {
      text-align: center;
      margin-top: 2rem;
      padding-top: 2rem;
      border-top: 1px solid var(--border-color);
      color: var(--text-secondary);
      font-size: 0.85rem;
    }

    .footer-text a {
      color: var(--accent-color);
      text-decoration: none;
      transition: color 0.3s ease;
    }

    .footer-text a:hover {
      color: var(--gold-accent);
    }

    /* Loading spinner */
    .button-spinner {
      display: inline-block;
      width: 14px;
      height: 14px;
      border: 2px solid rgba(255, 255, 255, 0.3);
      border-top: 2px solid var(--text-primary);
      border-radius: 50%;
      animation: spin 0.8s linear infinite;
      margin-right: 0.5rem;
    }

    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }

    /* Professional scrollbar */
    ::-webkit-scrollbar {
      width: 6px;
    }

    ::-webkit-scrollbar-track {
      background: var(--bg-secondary);
    }

    ::-webkit-scrollbar-thumb {
      background: var(--border-color);
      border-radius: 3px;
    }

    ::-webkit-scrollbar-thumb:hover {
      background: var(--accent-color);
    }

    @media (max-width: 480px) {
      .login-card {
        padding: 2rem 1.5rem;
      }

      .login-title {
        font-size: 1.5rem;
      }

      .logo-container {
        width: 70px;
        height: 70px;
      }

      .logo-container i {
        font-size: 2rem;
      }
    }
  </style>
</head>
<body>
  <!-- Animated background particles -->
  <div class="bg-particles" id="particles"></div>

  <div class="login-container">
    <div class="login-card">
      <div class="login-header">
        <div class="logo-container">
          <i class="fa-solid fa-money-check-dollar"></i>
        </div>
        <h1 class="login-title">BMS</h1>
        <p class="login-subtitle">Budget Management System</p>
      </div>

      <div class="alert-message" id="alertMessage">
        <i class="fas fa-exclamation-circle me-2"></i>
        <span id="alertText">Invalid credentials</span>
      </div>

      <!-- ASP.NET Form -->
      <form runat="server" id="form1">
        <asp:ScriptManager runat="server" />
        
        <div class="form-group">
          <label for="txtEmail" class="form-label">Email Address</label>
          <div class="input-wrapper">
            <asp:TextBox 
              runat="server" 
              ID="txtEmail" 
              CssClass="form-control" 
              placeholder="อีเมล / Email (@kingpower.com)"
              TextMode="Email"
            />
            <i class="fas fa-envelope input-icon"></i>
          </div>
        </div>

        <div class="form-group">
          <label for="txtPassword" class="form-label">Password</label>
          <div class="input-wrapper">
            <asp:TextBox 
              runat="server" 
              ID="txtPassword" 
              CssClass="form-control" 
              TextMode="Password" 
              placeholder="รหัสผ่าน / Password"
            />
            <i class="fas fa-lock input-icon"></i>
            <button 
              type="button" 
              class="password-toggle" 
              onclick="togglePassword()"
              aria-label="Toggle password visibility"
            >
              <i class="fas fa-eye" id="toggleIcon"></i>
            </button>
          </div>
        </div>

        <div class="remember-forgot">
          <div class="remember-me">
            <input type="checkbox" id="remember" name="remember">
            <label for="remember">Remember me</label>
          </div>
          <a href="#" class="forgot-link">Forgot Password?</a>
        </div>

        <asp:Button 
          runat="server" 
          ID="btnSubmit" 
          CssClass="btn-kp-primary w-100 mb-3" 
          Text="เข้าสู่ระบบ / Log in"
        />
      </form>

      <div class="footer-text">
        <p>© 2025 Powered by CIE. All rights reserved.</p>
        <p class="mt-2">
          <a href="#">Privacy Policy</a> • <a href="#">Terms of Service</a>
        </p>
      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
  
  <script>
      // Create animated particles
      function createParticles() {
          const container = document.getElementById('particles');
          const particleCount = 30;

          for (let i = 0; i < particleCount; i++) {
              const particle = document.createElement('div');
              particle.className = 'particle';
              particle.style.left = Math.random() * 100 + '%';
              particle.style.animationDuration = (Math.random() * 10 + 10) + 's';
              particle.style.animationDelay = Math.random() * 5 + 's';
              container.appendChild(particle);
          }
      }

      // Toggle password visibility
      function togglePassword() {
          // Get the ASP.NET generated password input
          const passwordInput = document.querySelector('input[type="password"][id*="txtPassword"]') ||
              document.getElementById('txtPassword');
          const toggleIcon = document.getElementById('toggleIcon');

          if (passwordInput) {
              if (passwordInput.type === 'password') {
                  passwordInput.type = 'text';
                  toggleIcon.classList.remove('fa-eye');
                  toggleIcon.classList.add('fa-eye-slash');
              } else {
                  passwordInput.type = 'password';
                  toggleIcon.classList.remove('fa-eye-slash');
                  toggleIcon.classList.add('fa-eye');
              }
          }
      }

      // Show alert message (can be called from code-behind)
      function showAlert(message, isError = true) {
          const alertMessage = document.getElementById('alertMessage');
          const alertText = document.getElementById('alertText');

          alertText.textContent = message;
          alertMessage.classList.add('show');

          if (!isError) {
              alertMessage.style.background = 'rgba(76, 175, 80, 0.1)';
              alertMessage.style.borderColor = 'rgba(76, 175, 80, 0.3)';
              alertMessage.style.color = '#4caf50';
          } else {
              alertMessage.style.background = 'rgba(244, 67, 54, 0.1)';
              alertMessage.style.borderColor = 'rgba(244, 67, 54, 0.3)';
              alertMessage.style.color = '#f44336';
          }

          setTimeout(() => {
              alertMessage.classList.remove('show');
          }, 5000);
      }

      // Initialize
      document.addEventListener('DOMContentLoaded', function () {
          createParticles();

          // Auto-focus email field
          const emailInput = document.querySelector('input[type="email"][id*="txtEmail"]') ||
              document.querySelector('input[id*="txtEmail"]');
          if (emailInput) {
              emailInput.focus();
          }
      });

      // Handle ASP.NET form submission (add loading state)
      if (typeof Sys !== 'undefined') {
          Sys.WebForms.PageRequestManager.getInstance().add_beginRequest(function () {
              const submitBtn = document.querySelector('input[id*="btnSubmit"]');
              if (submitBtn) {
                  submitBtn.disabled = true;
                  submitBtn.value = 'กำลังเข้าสู่ระบบ...';
              }
          });

          Sys.WebForms.PageRequestManager.getInstance().add_endRequest(function () {
              const submitBtn = document.querySelector('input[id*="btnSubmit"]');
              if (submitBtn) {
                  submitBtn.disabled = false;
                  submitBtn.value = 'เข้าสู่ระบบ / Log in';
              }
          });
      }
  </script>
</body>
</html>

