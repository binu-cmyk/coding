import subprocess
import time

# =========================
# CONFIG
# =========================
VBS_FILE = r"D:\Project Backup\VS Script for Gypsum QC.vbs"

# =========================
# DATES (same as your Excel)
# =========================
FROM_DATE = "01.04.2025"
TO_DATE   = "31.05.2026"

# =========================
# RUN SAP
# =========================
print("🚀 Starting SAP Automation...")

time.sleep(2)

subprocess.call([
    "cscript.exe",
    VBS_FILE,
    FROM_DATE,
    TO_DATE
])

print("✅ SAP extraction completed successfully.")