# WSUS Update Management Script - Enhanced Version

## Overview
This PowerShell script automates the approval and declination of updates in a **WSUS (Windows Server Update Services)** server. It builds upon the original script **`wsus-approval.ps1`**, acquired from [this GitHub repository](https://github.com/hkbakke/wsus-helpers), and introduces several key improvements.

### **Original Script Source:**
- **File Name:** `wsus-approval.ps1`
- **Repository:** [hkbakke/wsus-helpers](https://github.com/hkbakke/wsus-helpers)
- **Purpose:** Automate WSUS update approvals and declinations

## 🚀 Key Enhancements & New Features
### **1️⃣ Improved Documentation & Code Structure**
✅ Added **structured help comments** (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`, `.NOTES`)
✅ Introduced **debug values section** for easier testing
✅ Enhanced code readability with **better modularization**

### **2️⃣ New & Improved Parameters**
| Parameter | Description |
|------------|----------------|
| `-DeclineOnly` | Skips approvals, only declines updates |
| `-IncludeUpgrades` | Optionally includes upgrade classifications in approvals |
| `-RestrictToLanguages` | Filters updates based on specified locales (default: `en-us`, `en-gb`) |

### **3️⃣ Enhanced Logging System**
✅ **Timestamped logging** with **colored console output** for better readability
✅ Logs written to a **rotating log file** (`wsus-approver_Day_Hour.log`)
✅ More structured logging for **debugging and traceability**

### **4️⃣ Smarter Update Filtering & Approval Process**
✅ **New function:** `TestUpdateTitleLanguageMatch` - prevents non-English updates from being approved
✅ **Replaced inefficient `if-else` conditions** with a **PowerShell `switch` statement** for better performance
✅ **Declines superseded/expired updates only after approving new ones**
✅ **Handles license agreement acceptance more explicitly** before approving updates

### **5️⃣ More Robust Synchronization Handling**
✅ Ensures **WSUS synchronization is complete before processing updates**
✅ **Waits for ongoing syncs to finish** instead of blindly proceeding
✅ Prevents race conditions and **sync conflicts**

### **6️⃣ Optimized Execution & Error Handling**
✅ Sets **default error action to `Stop`** to ensure failures are caught
✅ **Prevents unnecessary processing** on deselected updates
✅ **More modular functions** for better maintenance

## 🛠️ Comparison Summary
| **Feature** | **Original (`wsus-approval.ps1`)** | **Updated Version** | **Benefit** |
|------------|---------------------------------|-----------------|-------------|
| **Logging** | Basic text logging | Timestamped, colored logging | Easier debugging & visibility |
| **Approval Process** | Fixed categories only | Supports **upgrades** (optional) | More flexibility |
| **Decline Process** | No language filtering | **Filters by locale** | Avoids unnecessary approvals |
| **Parameters** | Limited options | New `DeclineOnly`, `IncludeUpgrades`, `RestrictToLanguages` | More control |
| **Synchronization** | Immediate processing | Waits for sync completion | Prevents conflicts |
| **Code Readability** | Flat structure | Modularized with functions & docs | Easier maintenance |

## 🔧 How to Use
### **Basic Example:**
```powershell
# Connects to WSUS and runs a dry run without making any changes
.\WSUS-Approver.ps1 -WsusServer "wsus-server" -Port 8531 -UseSSL -DryRun
```

### **Advanced Example:**
```powershell
# Approve updates including upgrades, decline non-English updates, and perform real changes
.\WSUS-Approver.ps1 -IncludeUpgrades -RestrictToLanguages @("en-us") -DryRun:$false
```