<#
.SYNOPSIS
Simulates pressing Ctrl+Alt+Tab, then Tab, then Enter every 5 seconds
to cycle through ALL open application windows sequentially.

.DESCRIPTION
This script uses the Ctrl+Alt+Tab keyboard shortcut to open the persistent
Windows task switcher UI. It then sends a Tab keystroke to advance the selection
to the next window in the cycle, followed by an Enter keystroke to activate
that window. This process repeats every 5 seconds, effectively cycling through
all major open application windows one by one.

.NOTES
Author: Your Name/AI Assistant
Date:   2023-10-27
Requires: Windows PowerShell or PowerShell 7+
Warning: Using this script to bypass activity monitoring may violate company policy.
         Press CTRL+C in the PowerShell window to stop the script.
         The short delays (Milliseconds) might need adjustment on slower systems.
#>

# Load the necessary assembly to send keys
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Write-Host "System.Windows.Forms assembly loaded successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to load System.Windows.Forms assembly. This script requires it to send keystrokes."
    Write-Error "Error details: $($_.Exception.Message)"
    # Exit if the assembly can't be loaded
    exit 1
}

# --- Script Configuration ---
$MainIntervalSeconds = 5 # Time between window switches
$UISettleDelayMs = 250    # Milliseconds to wait after Ctrl+Alt+Tab and after Tab
# --------------------------

Write-Host "Starting window cycling simulation..." -ForegroundColor Cyan
Write-Host "Switching windows every $MainIntervalSeconds seconds using Ctrl+Alt+Tab -> Tab -> Enter." -ForegroundColor Cyan
Write-Host "Switch to this PowerShell window and press CTRL+C to stop the script." -ForegroundColor Yellow

try {
    # Infinite loop - runs until manually stopped (CTRL+C)
    while ($true) {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Ctrl+Alt+Tab..."
        # Send Ctrl+Alt+Tab (^ is Ctrl, % is Alt)
        [System.Windows.Forms.SendKeys]::SendWait("^%{TAB}")

        # Wait briefly for the switcher UI to appear and stabilize
        Start-Sleep -Milliseconds $UISettleDelayMs

        # Wait briefly for the selection highlight to move
        Start-Sleep -Milliseconds $UISettleDelayMs

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Enter..."
        # Send Enter to activate the selected window
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Switch complete. Waiting $MainIntervalSeconds seconds..."
        # Wait for the main interval before the next cycle
        Start-Sleep -Seconds $MainIntervalSeconds


        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Ctrl+Alt+Tab..."
        # Send Ctrl+Alt+Tab (^ is Ctrl, % is Alt)
        [System.Windows.Forms.SendKeys]::SendWait("^%{TAB}")

        # Wait briefly for the switcher UI to appear and stabilize
        Start-Sleep -Milliseconds $UISettleDelayMs

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Tab..."
        # Send Tab to move selection to the next window
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

        # Wait briefly for the selection highlight to move
        Start-Sleep -Milliseconds $UISettleDelayMs

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Enter..."
        # Send Enter to activate the selected window
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Switch complete. Waiting $MainIntervalSeconds seconds..."
        # Wait for the main interval before the next cycle
        Start-Sleep -Seconds $MainIntervalSeconds


        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Ctrl+Alt+Tab..."
        # Send Ctrl+Alt+Tab (^ is Ctrl, % is Alt)
        [System.Windows.Forms.SendKeys]::SendWait("^%{TAB}")

        # Wait briefly for the switcher UI to appear and stabilize
        Start-Sleep -Milliseconds $UISettleDelayMs

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Tab..."
        # Send Tab to move selection to the next window
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Tab..."
        # Send Tab to move selection to the next window
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")


        # Wait briefly for the selection highlight to move
        Start-Sleep -Milliseconds $UISettleDelayMs

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Enter..."
        # Send Enter to activate the selected window
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Switch complete. Waiting $MainIntervalSeconds seconds..."
        # Wait for the main interval before the next cycle
        Start-Sleep -Seconds $MainIntervalSeconds




        
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Ctrl+Alt+Tab..."
        # Send Ctrl+Alt+Tab (^ is Ctrl, % is Alt)
        [System.Windows.Forms.SendKeys]::SendWait("^%{TAB}")

        # Wait briefly for the switcher UI to appear and stabilize
        Start-Sleep -Milliseconds $UISettleDelayMs

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Tab..."
        # Send Tab to move selection to the next window
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Tab..."
        # Send Tab to move selection to the next window
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

         Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Tab..."
        # Send Tab to move selection to the next window
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")


        # Wait briefly for the selection highlight to move
        Start-Sleep -Milliseconds $UISettleDelayMs

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Enter..."
        # Send Enter to activate the selected window
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Switch complete. Waiting $MainIntervalSeconds seconds..."
        # Wait for the main interval before the next cycle
        Start-Sleep -Seconds $MainIntervalSeconds




         Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Ctrl+Alt+Tab..."
        # Send Ctrl+Alt+Tab (^ is Ctrl, % is Alt)
        [System.Windows.Forms.SendKeys]::SendWait("^%{TAB}")

        # Wait briefly for the switcher UI to appear and stabilize
        Start-Sleep -Milliseconds $UISettleDelayMs

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Tab..."
        # Send Tab to move selection to the next window
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Tab..."
        # Send Tab to move selection to the next window
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Tab..."
        # Send Tab to move selection to the next window
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Tab..."
        # Send Tab to move selection to the next window
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

        # Wait briefly for the selection highlight to move
        Start-Sleep -Milliseconds $UISettleDelayMs

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Sending Enter..."
        # Send Enter to activate the selected window
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Switch complete. Waiting $MainIntervalSeconds seconds..."
        # Wait for the main interval before the next cycle
        Start-Sleep -Seconds $MainIntervalSeconds
    }
}
catch [System.Management.Automation.PipelineStoppedException] {
    # Catch the exception thrown when CTRL+C is pressed
    Write-Host "`nScript stopped by user (CTRL+C)." -ForegroundColor Green
}
catch {
    # Catch any other unexpected errors
    Write-Error "An unexpected error occurred: $($_.Exception.Message)"
}
finally {
    Write-Host "Exiting script."
}