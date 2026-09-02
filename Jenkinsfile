pipeline {
    agent none 

    options {
        disableConcurrentBuilds()
        timestamps()
        buildDiscarder(logRotator(numToKeepStr: '10'))
    }

    parameters {
        choice(name: 'ENVIRONMENT', choices: ['UAT', 'PROD'], description: 'เลือก Environment ที่ต้องการ Deploy')
        choice(name: 'ROLLBACK_VERSION', choices: [
            'None', 
            'Latest (1)', 
            'Previous (2)', 
            'Older (3)', 
            'Older (4)', 
            'Oldest (5)'
        ], description: 'เลือกเวอร์ชันที่ต้องการ Rollback (None = ไม่ Rollback, Deploy โค้ดใหม่ตามปกติ)')
    }

    environment {
        IIS_SITE_PATH = "E:\\www\\BMS_Test" 
        IIS_BACKUP_PATH = "E:\\www\\BMS_backup" 
        APP_POOL_NAME = "BMS_Test" 
        TARGET_NODE = "${params.ENVIRONMENT == 'PROD' ? 'otb-prod' : 'otb-uat'}"
        APP_URL = "${params.ENVIRONMENT == 'PROD' ? 'https://bms.kingpower.com' : 'https://dev-cie.kingpower.com/bms'}" 
    }

    stages {
        stage('Approval PROD') {
            when { expression { params.ENVIRONMENT == 'PROD' } }
            steps {
                timeout(time: 15, unit: 'MINUTES') {
                    input message: '⚠️ คุณกำลังจะแก้ไขระบบ Production (PROD) ยืนยันหรือไม่?', ok: '✅ ยืนยันการ Deploy'
                }
            }
        }

        // ===================================================
        // 1. SOURCE & COMPILE
        // ===================================================
        stage('Source & Compile') {
            when { 
                beforeAgent true 
                expression { params.ROLLBACK_VERSION == 'None' } 
            }
            agent { label 'built-in' }
            stages {
                stage('Checkout') {
                    steps {
                        checkout scm
                    }
                }
                stage('Build') {
                    steps {
                        echo "Restoring NuGet packages..."
                        bat 'nuget restore ./BMS.vbproj -SolutionDirectory . -MSBuildPath "C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\MSBuild\\Current\\Bin"'
                        
                        bat '"C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\MSBuild\\Current\\Bin\\MSBuild.exe" ./BMS.vbproj /p:Configuration=Release /p:DeployOnBuild=true /p:WebPublishMethod=FileSystem'
                        
                        stash includes: 'obj/Release/Package/PackageTmp/**', name: 'compiled-app'
                    }
                }
            }
            post {
                always {
                    cleanWs() 
                }
            }
        }

        // ===================================================
        // 2. BACKUP STAGE
        // ===================================================
        stage('Backup System') {
            when { expression { params.ROLLBACK_VERSION == 'None' } }
            steps {
                script {
                    node(env.TARGET_NODE) { 
                        try {
                            powershell '''
                            $TIMESTAMP = Get-Date -Format "yyyyMMdd_HHmmss"
                            $TARGET_BACKUP = "$env:IIS_BACKUP_PATH\\$TIMESTAMP"
                            
                            Write-Host "Creating backup at $TARGET_BACKUP..."
                            if (!(Test-Path $env:IIS_BACKUP_PATH)) {
                                New-Item -ItemType Directory -Force -Path $env:IIS_BACKUP_PATH | Out-Null
                            }
                            
                            New-Item -ItemType Directory -Force -Path $TARGET_BACKUP | Out-Null
                            if (Test-Path $env:IIS_SITE_PATH) {
                                Copy-Item -Path "$env:IIS_SITE_PATH\\*" -Destination $TARGET_BACKUP -Recurse -Force
                            }
                            
                            Write-Host "Cleaning up old backups (Keeping last 5)..."
                            Get-ChildItem -Path $env:IIS_BACKUP_PATH -Directory | Sort-Object CreationTime -Descending | Select-Object -Skip 5 | Remove-Item -Recurse -Force
                            '''
                        } finally {
                            cleanWs()
                        }
                    }
                }
            }
        }

        // ===================================================
        // 3. DEPLOY TO IIS (ก๊อปปี้ทุกไฟล์ ยกเว้น Web.config ไม่ให้ทับของเดิม)
        // ===================================================
        stage('Deploy to IIS') {
            when { expression { params.ROLLBACK_VERSION == 'None' } }
            steps {
                script {
                    node(env.TARGET_NODE) { 
                        try {
                            powershell '''
                            Write-Host "Stopping IIS Application Pool..."
                            Import-Module WebAdministration
                            Stop-WebAppPool -Name $env:APP_POOL_NAME -ErrorAction SilentlyContinue
                            Start-Sleep -Seconds 3
                            
                            # ลบไฟล์เก่าทั้งหมด ยกเว้น Web.config ตัวจริงบนเซิร์ฟเวอร์
                            Get-ChildItem -Path $env:IIS_SITE_PATH -Exclude "Web.config" | Remove-Item -Recurse -Force
                            '''
                            
                            unstash 'compiled-app'
                            
                            powershell '''
                            Write-Host "Copying new compiled files to IIS (Excluding Web.config)..."
                            
                            # ก๊อปปี้ไฟล์ใหม่ทั้งหมด ยกเว้น Web.config ไม่ให้นำมาทับของเดิม
                            Get-ChildItem -Path ".\\obj\\Release\\Package\\PackageTmp" -Recurse | Where-Object { $_.Name -ne "Web.config" } | ForEach-Object {
                                $relativePath = $_.FullName.Substring((Resolve-Path ".\\obj\\Release\\Package\\PackageTmp").Path.Length)
                                $destPath = Join-Path $env:IIS_SITE_PATH $relativePath
                                
                                if ($_.PSIsContainer) {
                                    if (!(Test-Path $destPath)) { New-Item -ItemType Directory -Force -Path $destPath | Out-Null }
                                } else {
                                    $parentDir = Split-Path $destPath -Parent
                                    if (!(Test-Path $parentDir)) { New-Item -ItemType Directory -Force -Path $parentDir | Out-Null }
                                    Copy-Item -Path $_.FullName -Destination $destPath -Force
                                }
                            }
                            
                            Start-WebAppPool -Name $env:APP_POOL_NAME
                            Write-Host "✅ New files deployed (Original Web.config preserved) and Application Pool started."
                            '''
                        } finally {
                            cleanWs()
                        }
                    }
                }
            }
        }

        // ===================================================
        // 4. ROLLBACK SYSTEM
        // ===================================================
        stage('Rollback System') {
            when { expression { params.ROLLBACK_VERSION != 'None' } }
            steps {
                script {
                    node(env.TARGET_NODE) { 
                        try {
                            withEnv(["ROLLBACK_CHOICE=${params.ROLLBACK_VERSION}"]) {
                                powershell '''
                                Write-Host "Initiating Manual Rollback on $env:ENVIRONMENT..."
                                
                                if (Test-Path $env:IIS_BACKUP_PATH) {
                                    $skipCount = 0
                                    if ($env:ROLLBACK_CHOICE -match "1") { $skipCount = 0 }
                                    elseif ($env:ROLLBACK_CHOICE -match "2") { $skipCount = 1 }
                                    elseif ($env:ROLLBACK_CHOICE -match "3") { $skipCount = 2 }
                                    elseif ($env:ROLLBACK_CHOICE -match "4") { $skipCount = 3 }
                                    elseif ($env:ROLLBACK_CHOICE -match "5") { $skipCount = 4 }
                                    
                                    $SELECTED_BACKUP = Get-ChildItem -Path $env:IIS_BACKUP_PATH -Directory | Sort-Object CreationTime -Descending | Select-Object -Skip $skipCount -First 1
                                    
                                    if ($SELECTED_BACKUP) {
                                        Write-Host "Restoring from $($SELECTED_BACKUP.FullName)..."
                                        Import-Module WebAdministration
                                        Stop-WebAppPool -Name $env:APP_POOL_NAME -ErrorAction SilentlyContinue
                                        Start-Sleep -Seconds 2
                                        
                                        Remove-Item -Path "$env:IIS_SITE_PATH\\*" -Recurse -Force
                                        Copy-Item -Path "$($SELECTED_BACKUP.FullName)\\*" -Destination $env:IIS_SITE_PATH -Recurse -Force
                                        
                                        Start-WebAppPool -Name $env:APP_POOL_NAME
                                        Write-Host "✅ Manual Rollback Completed Successfully."
                                    } else {
                                        Write-Error "❌ Backup version '$env:ROLLBACK_CHOICE' not found!"
                                        exit 1
                                    }
                                } else {
                                    Write-Error "❌ Backup path does not exist!"
                                    exit 1
                                }
                                '''
                            }
                        } finally {
                            cleanWs()
                        }
                    }
                }
            }
        }

        // ===================================================
        // 5. HEALTH CHECK
        // ===================================================
        stage('Health Check') {
            steps {
                script {
                    node(env.TARGET_NODE) { 
                        try {
                            powershell '''
                            Write-Host "Waiting 5 seconds for application pool to spin up..."
                            Start-Sleep -Seconds 5
                            
                            Write-Host "Performing Health Check on $env:APP_URL ..."
                            try {
                                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                                $response = Invoke-WebRequest -Uri $env:APP_URL -UseBasicParsing -TimeoutSec 30
                                if ($response.StatusCode -eq 200) {
                                    Write-Host "✅ Health Check Passed! Status: $($response.StatusCode)"
                                } else {
                                    Write-Error "❌ Health Check Failed! Status: $($response.StatusCode)"
                                    exit 1
                                }
                            } catch {
                                Write-Error "❌ Health Check Exception: Application might be crashing. Error details: $($_.Exception.Message)"
                                exit 1
                            }
                            '''
                        } finally {
                            cleanWs()
                        }
                    }
                }
            }
        }

        // ===================================================
        // 6. SUCCESS STAGE
        // ===================================================
        stage('Deploy Success') {
            steps {
                echo "✅ Pipeline finished successfully! (All stages passed)"
            }
        }
    }

    // ===================================================
    // POST ACTIONS
    // ===================================================
    post {
        failure {
            script {
                if (params.ROLLBACK_VERSION == 'None') {
                    node(env.TARGET_NODE) {
                        powershell '''
                        if ((Test-Path $env:IIS_BACKUP_PATH) -and ((Get-ChildItem -Path $env:IIS_BACKUP_PATH -Directory).Count -gt 0)) {
                            Write-Host "❌ Pipeline failed! Initiating Auto-Rollback..."
                            $LATEST_BACKUP = Get-ChildItem -Path $env:IIS_BACKUP_PATH -Directory | Sort-Object CreationTime -Descending | Select-Object -First 1
                            if ($LATEST_BACKUP) {
                                Write-Host "Restoring from $($LATEST_BACKUP.FullName)..."
                                Import-Module WebAdministration
                                Stop-WebAppPool -Name $env:APP_POOL_NAME -ErrorAction SilentlyContinue
                                Start-Sleep -Seconds 2
                                
                                if (Test-Path $env:IIS_SITE_PATH) {
                                    Remove-Item -Path "$env:IIS_SITE_PATH\\*" -Recurse -Force
                                }
                                Copy-Item -Path "$($LATEST_BACKUP.FullName)\\*" -Destination $env:IIS_SITE_PATH -Recurse -Force
                                
                                Start-WebAppPool -Name $env:APP_POOL_NAME
                                Write-Host "✅ Auto-Rollback Completed Successfully."
                            }
                        } else {
                            Write-Host "⚠️ Build failed before deployment stage. No existing backup restoration needed."
                        }
                        '''
                    }
                } else {
                    echo "❌ Pipeline failed during Manual Rollback."
                }
            }
        }
    }
}
