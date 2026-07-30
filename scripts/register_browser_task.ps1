# register_browser_task.ps1
# Registers a monthly Task Scheduler job to auto-download Claims Planner boards.
# Run once — no admin needed.
#
# The task runs at 6:00 AM on the 1st of each month.
#
# Usage:
#   .\scripts\register_browser_task.ps1

$ErrorActionPreference = "Stop"

# Self-locate
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ProjectDir = Split-Path -Parent $ScriptDir
$PythonExe  = Join-Path $ProjectDir ".venv\Scripts\python.exe"
$Script     = Join-Path $ScriptDir  "download_planner_browser.py"
$TaskName   = "Claims_Planner_Download"
$LogDir     = Join-Path $ProjectDir "data\logs"
$LogFile    = Join-Path $LogDir "browser_download.log"

# Verify venv exists
if (-not (Test-Path $PythonExe)) {
    Write-Warning "Venv not found at $PythonExe"
    Write-Warning "Build it:  python -m venv .venv && .venv\Scripts\pip install -r requirements.txt && .venv\Scripts\playwright install chromium"
}

# Ensure log directory exists
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Force $LogDir | Out-Null }

# Remove existing task silently
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

# Build task XML — avoids all schtasks.exe quoting issues with spaces in paths.
# Monthly CalendarTrigger: 1st of every month at 06:00.
# Logging: wrap in cmd /c so >> redirect works.
$CmdArgs = "/c `"`"$PythonExe`" `"$Script`" >> `"$LogFile`" 2>&1`""

$TaskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Download Claims Planner boards (Ford + Chrysler) monthly.</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>2026-07-01T06:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByMonth>
        <DaysOfMonth><Day>1</Day></DaysOfMonth>
        <Months>
          <January/><February/><March/><April/>
          <May/><June/><July/><August/>
          <September/><October/><November/><December/>
        </Months>
      </ScheduleByMonth>
    </CalendarTrigger>
  </Triggers>
  <Settings>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <StartWhenAvailable>true</StartWhenAvailable>
    <Enabled>true</Enabled>
  </Settings>
  <Principals>
    <Principal id="Author">
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Actions Context="Author">
    <Exec>
      <Command>cmd.exe</Command>
      <Arguments>$CmdArgs</Arguments>
      <WorkingDirectory>$ProjectDir</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

Register-ScheduledTask -TaskName $TaskName -Xml $TaskXml -Force | Out-Null

Write-Host "Registered: $TaskName"
Write-Host "  Schedule : 1st of each month at 06:00 AM"
Write-Host "  Script   : $Script"
Write-Host "  Python   : $PythonExe"
Write-Host "  Log      : $LogFile"
Write-Host ""
Write-Host "To test immediately:"
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host ""
Write-Host "To verify:"
Write-Host "  Get-ScheduledTask -TaskName '$TaskName' | Select TaskName,State"
