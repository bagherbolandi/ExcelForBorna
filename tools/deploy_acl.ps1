# ============================================================================
#  Borna Composite — شبکه‌سازی دسترسی پروژه‌ها (icacls)
#  اجرا: PowerShell با حقوق Administrator روی سرور فایل
#  پیش‌نیاز: گروه‌های AD طبق docs/IT-DEPLOYMENT.md ساخته شده باشند.
# ============================================================================

$ErrorActionPreference = "Stop"

# --- پیکربندی: فقط این بخش را عوض کنید --------------------------------------
$Root        = "C:\Shares\BornaProjects"                 # ریشه Share
$ShareName   = "BornaProjects$"
$DomainFQDN  = "CORP.BORNACOMPPOSITE.LOCAL"              # دامنه ویندوز شرکت

# گروه‌های AD (نام کوتاه) → سطح دسترسی
$Groups = [ordered]@{
    "Borna-PMO"     = "Modify"
    "Borna-IT"      = "FullControl"
}
$ReadOnly = @("Borna-Exec")

# هر پوشه واحد → گروهی که Modify می‌گیرد
$FolderRights = [ordered]@{
    "01_PM"   = @("Borna-PMO")
    "02_ENG"  = @("Borna-Eng")
    "03_TRD"  = @("Borna-TradeIn","Borna-TradeEx")
    "04_PLAN" = @("Borna-Plan")
    "05_FIN"  = @("Borna-Fin")
    "06_SALES"= @("Borna-Sales")
    "07_EXEC" = @("Borna-PMO")
    "08_RISK" = @("Borna-PMO","Borna-IT")
}
$Project = "BC-P-1405-01_FRP-Tank"    # نام پوشه پروژه (برای هر پروژه یک‌بار اجرا/ویرایش)

# --- کمکی --------------------------------------------------------------------
function Grant-NTFS {
    param([string]$Path, [string]$Group, [string]$Level)
    $map = @{
        "Modify"      = "(OI)(CI)M"
        "Read"        = "(OI)(CI)RX"
        "FullControl" = "(OI)(CI)F"
    }
    $sid = (New-Object System.Security.Principal.NTAccount("$DomainFQDN\$Group")).Translate([System.Security.Principal.SecurityIdentifier]).Value
    icacls "`"$Path`"" /inheritance:r /grant "*${sid}:${map[$Level]}" | Out-Null
    Write-Host "  $Group → $Level  on $Path"
}
function Add-SpecialDenied {
    # (اختیاری) — در نسخهٔ پیش‌فرض استفاده نمی‌شود
    # حذف امکان تغییر نام/مالکیت برای غیرادمین‌ها روی فایل قفل‌شده
    param([string]$Path, [string]$Group)
    $sid = (New-Object System.Security.Principal.NTAccount("$DomainFQDN\$Group")).Translate([System.Security.Principal.SecurityIdentifier]).Value
    icacls "`"$Path`"" /deny "*${sid}:(DC,WO,WDAC)" | Out-Null
}

# --- ساختار ریشه --------------------------------------------------------------
New-Item -ItemType Directory -Force -Path "$Root\_TEMPLATE", "$Root\_LOGS" | Out-Null
Grant-NTFS "$Root\_TEMPLATE" "Borna-PMO" "Read"
Grant-NTFS "$Root\_TEMPLATE" "Borna-IT"  "FullControl"
foreach ($g in $ReadOnly) { Grant-NTFS "$Root\_TEMPLATE" $g "Read" }
Grant-NTFS "$Root\_LOGS" "Borna-IT" "FullControl"

# --- پوشه پروژه ---------------------------------------------------------------
# IMPORTANT / نکته معماری:
#  چون کل پروژه یک فایل است، هر واحدی که در فایلی می‌نویسد باید روی آن
#  Modify (فایل‌سیستم) داشته باشد؛ «هر واحد فقط شیت خودش» در اکسل فقط در سطح
#  UI ممکن است (VBA + Sheet Protection). برای جداسازی سخت‌افزاری داده،
#  حالت چندفایلی (بخش ۴ سند IT) را اجرا کنید.
$proj = Join-Path $Root $Project
New-Item -ItemType Directory -Force -Path $proj | Out-Null
$AllWriters = ($FolderRights.Values | ForEach-Object { $_ }) | Sort-Object -Unique

foreach ($sub in $FolderRights.Keys) {
    $p = Join-Path $proj $sub
    New-Item -ItemType Directory -Force -Path $p | Out-Null
    Grant-NTFS $p "Borna-PMO" "Modify"
    Grant-NTFS $p "Borna-IT"  "FullControl"
    foreach ($g in $ReadOnly) { Grant-NTFS $p $g "Read" }
    foreach ($g in $FolderRights[$sub]) { Grant-NTFS $p $g "Modify" }
    # بقیه واحدها: Read روی پوشه‌های دیگر (بدون تغییر فایل‌ها)
    foreach ($g in $AllWriters) {
        if ($g -notin $FolderRights[$sub]) { Grant-NTFS $p $g "Read" }
    }
}
# فایل اصلی: تیم پروژه Modify (اجازهٔ ذخیرهٔ اکسل) — مدیریت ارشد فقط Read
Grant-NTFS $proj "Borna-PMO" "Modify"
Grant-NTFS $proj "Borna-IT"  "FullControl"
foreach ($g in $ReadOnly) { Grant-NTFS $proj $g "Read" }
foreach ($g in $AllWriters) {
    if ($g -notin @("Borna-PMO","Borna-IT")) { Grant-NTFS $proj $g "Modify" }
}
# قفل ساختار فایل: تغییر نام/حذف فقط توسط مالک (PMO) — از طریق Owner
takeown /f "$proj\Project.xlsm" /a | Out-Null

Write-Host "`nACL برای $proj اعمال شد. تأیید: icacls `"$proj`"" -ForegroundColor Green
Write-Host "گام بعدی: فایل الگو را در $Root\_TEMPLATE بگذارید و برای پروژه کپی کنید (پاورپوینت‌نویسی نشده؛ دستی)." -ForegroundColor Yellow
