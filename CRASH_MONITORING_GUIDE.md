# How to Monitor AudioSimulation.exe Crash Logs

## Problem

Windows Event Viewer doesn't show application crash logs because:
1. App catches exceptions but doesn't report to Windows WER
2. Crash occurs in unmanaged code (C++/Native)
3. .NET Runtime errors not logged
4. Insufficient permissions

---

## Solution A: Basic PowerShell Monitor (Recommended - Simplest)

**Files**: `monitor_crash.ps1` + `start_monitor.bat`

**Usage**:
```bash
# Method 1: Double-click
start_monitor.bat

# Method 2: Run from PowerShell
powershell -ExecutionPolicy Bypass -File monitor_crash.ps1
```

**Monitors**:
- Process disappearance detection (crash exit)
- Application Error events
- Process restart detection
- Real-time log output to file

**Pros**: No extra tools needed, works immediately

---

## Solution B: Advanced ProcDump Monitor (Most Powerful)

**File**: `monitor_crash_advanced.ps1`

**Usage**:
```bash
powershell -ExecutionPolicy Bypass -File monitor_crash_advanced.ps1
```

**First run**: Automatically downloads Microsoft's official ProcDump tool

**Monitors**:
- Full memory dumps (.dmp files) for debugging
- Captures all unhandled exceptions
- Parallel Windows event log monitoring
- WER (Windows Error Reporting) monitoring
- System info logging

**Pros**: 
- Generates complete memory dump on crash, analyzable with WinDbg/Visual Studio
- Captures detailed exception info (type, call stack)

---

## Quick Start Guide

### Step 1: Choose Solution

- **Daily monitoring**: Use Solution A (`start_monitor.bat`)
- **Debugging crashes**: Use Solution B (`monitor_crash_advanced.ps1`)

### Step 2: Run Monitor

```bash
# Solution A
start_monitor.bat

# Solution B (requires admin for ProcDump)
# Right-click PowerShell -> Run as Administrator
powershell -ExecutionPolicy Bypass -File monitor_crash_advanced.ps1
```

### Step 3: Trigger App Crash

Run your `AudioSimulation.exe` and wait for or trigger crash scenarios

### Step 4: View Logs

**Solution A log location**:
```
E:\xk\Release\Release\crash_logs\crash_monitor_YYYYMMDD_HHMMSS.log
```

**Solution B log location**:
```
E:\xk\Release\Release\crash_logs\advanced_monitor_YYYYMMDD_HHMMSS.log
E:\xk\Release\Release\crash_dumps\AudioSimulation_crash_*.dmp
```

---

## Log Examples

### Process Crash Detection
```
[2026/4/30 14:25:33] Process disappeared! Likely crashed
[2026/4/30 14:25:35] Process restarted! New PID: 12345
```

### Application Error Event
```
[2026/4/30 14:25:33] Application error detected:
Faulting application name: AudioSimulation.exe
Version: 1.0.0.0
Timestamp: 0x647f8a2b
Exception code: 0xc0000005
Fault offset: 0x00012345
```

---

## Alternative Solutions

### Solution C: DebugView (Sysinternals)

Good for capturing OutputDebugString output:

1. Download: https://docs.microsoft.com/sysinternals/downloads/debugview
2. Run as Administrator
3. Check "Capture Global Win32" and "Capture Events"
4. Run app, view real-time debug output

### Solution D: Enable .NET CLR Logging

If app is .NET program:

```bash
# Enable .NET Runtime logging
reg add "HKLM\SOFTWARE\Microsoft\.NETFramework" /v EnableLog /t REG_DWORD /d 1 /f

# Enable Fusion logs (assembly binding)
reg add "HKLM\SOFTWARE\Microsoft\Fusion" /v ForceLog /t REG_DWORD /d 1 /f
reg add "HKLM\SOFTWARE\Microsoft\Fusion" /v LogPath /t REG_SZ /d "C:\FusionLogs" /f
```

### Solution E: PerfMon Performance Monitor

```bash
# Create data collector
perfmon /report
```

Add counters:
- Process > % Processor Time (AudioSimulation)
- Process > Handle Count (AudioSimulation)
- .NET CLR Exceptions > # of Exceps Thrown / sec

---

## Troubleshooting Tips

### If no logs appear

Possible causes:
1. **Silent crash** - Process killed by Task Manager/Antivirus
   - Check antivirus logs
   - Check Windows Defender quarantine

2. **Permission issues**
   - Run monitor script as Administrator
   - Check if app has write permissions

3. **Fast crash** (crashes on startup)
   - Use Solution B with `-w` parameter
   - Or use ProcDump directly:
     ```bash
     procdump.exe -e -ma -w AudioSimulation.exe crash_dump
     ```

4. **Child process crash**
   - Monitor all processes containing "AudioSimulation"
   - Modify script process name matching logic

---

## Recommended Complete Strategy

```
Daily Operation:
  └─ Solution A (run start_monitor.bat in background)

Debug Phase:
  └─ Solution B (ProcDump captures dump)
      + DebugView (capture debug output)
      + Enable .NET CLR logs

Production Environment:
  └─ Integrate NLog/log4net into app
      + Send crash reports to server
      + Local log rotation
```

---

## Contact Support

If above solutions can't capture crashes:
1. Add detailed try-catch logging in the app
2. Use third-party crash reporting service (Sentry, Bugsnag, etc.)
3. Contact app developer for debug symbol files (.pdb)
