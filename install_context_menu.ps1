# Додає / прибирає пункт "Вставити картинки в ексель" у контекстному меню Провідника.
# Пишеться тільки в HKCU — права адміністратора не потрібні.
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$key = 'YoInsertImgToExcel'
$title = 'Вставити картинки в ексель'
$bat = Join-Path $here 'insert_img_to_excel.bat'

# клік по папці -> %1, клік по порожньому місцю всередині папки -> %V
$targets = @{
    'HKCU:\Software\Classes\Directory\shell'            = '%1'
    'HKCU:\Software\Classes\Directory\Background\shell' = '%V'
}

foreach ($base in $targets.Keys) {
    $root = Join-Path $base $key
    if (Test-Path -LiteralPath $root) {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

if ($Uninstall) {
    Write-Host 'Пункт меню видалено.'
    exit 0
}

foreach ($base in $targets.Keys) {
    $arg = $targets[$base]
    $root = Join-Path $base $key
    New-Item -Path "$root\command" -Force | Out-Null
    New-ItemProperty -LiteralPath $root -Name 'MUIVerb' -Value $title -Force | Out-Null
    New-ItemProperty -LiteralPath $root -Name 'Icon' -Value "$env:SystemRoot\System32\imageres.dll,-5356" -Force | Out-Null
    Set-Item -LiteralPath "$root\command" -Value ('"{0}" "{1}"' -f $bat, $arg)
}

Write-Host "Готово. Клацніть правою кнопкою по папці -> `"$title`"."
Write-Host '(Windows 11: спершу "Показати додаткові параметри" / Shift+F10.)'
Write-Host "Скрипти запускаються з: $here"
