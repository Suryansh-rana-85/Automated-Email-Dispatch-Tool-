# =============================================================================
# PowerShell Script to Group Excel Data by Employee and send a personalized
# email with an attachment containing all of their specific rows.
# =============================================================================

try {
    # --- 1. Define File Paths and Read Configuration ---
    $scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
    $configPath = Join-Path $scriptDirectory "config.properties"

    if (-not (Test-Path $configPath)) {
        throw "Configuration file not found at: $configPath"
    }

    $config = @{}
    Get-Content $configPath | ForEach-Object {
        $line = $_.Trim()
        if ($line -and -not $line.StartsWith("#") -and $line.Contains("=")) {
            $parts = $line.Split('=', 2)
            $config[$parts[0].Trim()] = $parts[1].Trim()
        }
    }
    Write-Host "Configuration loaded successfully."

    # --- 2. Create a Temporary Folder for Attachments ---
    $tempFolderPath = Join-Path $scriptDirectory "temp_attachments"
    if (-not (Test-Path $tempFolderPath)) {
        New-Item -ItemType Directory -Path $tempFolderPath | Out-Null
        Write-Host "Created temporary attachment folder at: $tempFolderPath"
    }

    # --- 3. Ensure ImportExcel Module is Installed ---
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "The 'ImportExcel' module is not installed. Installing it now..."
        Install-Module -Name ImportExcel -Scope CurrentUser -Repository PSGallery -Force
    }

    # --- 4. Load and Group Excel Data ---
    $excelPath = Join-Path $scriptDirectory $config['mail.attachment.path']
    if (-not (Test-Path $excelPath)) {
        throw "Excel file not found at: $excelPath"
    }
    $allRecords = Import-Excel -Path $excelPath
    Write-Host "Successfully loaded $($allRecords.Count) total records from $($excelPath)."

    # Group all records by the 'Employee ID' column.
    $groupedByEmployee = $allRecords | Group-Object -Property 'Employee ID'
    Write-Host "Found $($groupedByEmployee.Count) unique employees to process." -ForegroundColor Yellow

    # --- 5. Prepare Email Credentials (done once) ---
    if ([string]::IsNullOrWhiteSpace($config['mail.username']) -or [string]::IsNullOrWhiteSpace($config['mail.password'])) {
        throw "The 'mail.username' or 'mail.password' is missing from your config.properties file."
    }
    $password = $config['mail.password'] | ConvertTo-SecureString -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($config['mail.username'], $password)
    
    # --- 6. Loop Through Each GROUP of Employee Records and Send Email ---
    $emailCounter = 0
    foreach ($employeeGroup in $groupedByEmployee) {
        
        # Get consistent details (like name and email) from the first record in the group.
        $firstRecord = $employeeGroup.Group[0]
        $recipientEmail = $firstRecord.Email
        $recipientName = $firstRecord.'First Name'

        if ([string]::IsNullOrWhiteSpace($recipientEmail)) {
            Write-Warning "Skipping group for Employee ID $($employeeGroup.Name) due to missing email."
            continue
        }
        
        Write-Host "Processing group for $recipientName (ID: $($employeeGroup.Name))..."

        # Create a Personalized Excel File Containing ALL Rows for this Employee
        $safeName = ($firstRecord.'First Name' + "_" + $firstRecord.'Last Name') -replace '[^a-zA-Z0-9_]', ''
        $tempAttachmentPath = Join-Path $tempFolderPath "Data_Summary_$($safeName).xlsx"
        
        # Export the ENTIRE group of rows ($employeeGroup.Group) to the Excel file.
        $employeeGroup.Group | Export-Excel -Path $tempAttachmentPath -AutoSize -TableName "YourDataSummary" -WorksheetName "Data"
        Write-Host "Generated Excel file with $($employeeGroup.Count) rows."

        # Build Personalized HTML Body
        $htmlBody = @"
        <div style='font-family: Arial, sans-serif; font-size: 14px;'>
            <h2 style='color: #005A9C;'>Confidential Data Summary</h2>
            <p>Hello $recipientName,</p>
            <p>Please find a summary of all your data records attached to this email.</p>
            <p>The attached file contains $($employeeGroup.Count) entries for your review.</p>
            <p>If you have any questions, please contact your administrator.</p>
        </div>
"@

        # Prepare Email and Send
        $emailParams = @{
            To          = $recipientEmail
            From        = $config['mail.username']
            Subject     = "$($config['mail.subject']) for $($recipientName)"
            Body        = $htmlBody
            BodyAsHtml  = $true
            SmtpServer  = $config['mail.smtp.host']
            Port        = $config['mail.smtp.port']
            UseSsl      = ($config['mail.smtp.starttls.enable'] -eq 'true')
            Credential  = $credential
            Attachments = $tempAttachmentPath
        }

        Send-MailMessage @emailParams
        $emailCounter++
        Write-Host "Email #$emailCounter sent successfully to $recipientEmail."

        # Clean up the temporary file
        Remove-Item -Path $tempAttachmentPath -Force
        
        # Wait Before Sending the Next Email
        $delay = [int]$config['mail.sendDelaySeconds']
        Write-Host "Waiting for $delay seconds..."
        Start-Sleep -Seconds $delay
    }

    # --- 7. Final Cleanup ---
    Remove-Item -Path $tempFolderPath -Recurse -Force
    Write-Host "Removed temporary attachment folder."
    
    Write-Host "-------------------------------------------------"
    Write-Host "Automation complete. Total emails sent to unique employees: $emailCounter." -ForegroundColor Green

} catch {
    Write-Error "A critical error occurred: $($_.Exception.Message)"
    Start-Sleep -Seconds 10
}