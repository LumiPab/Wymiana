<#
.Synopsis
Skrypt zbiera z komputera informacje pomocne przy procesie wymiany komputera

.DESCRIPTION
Skrypt ma za zadanie zebrać i wyśiwetlić użytkownikowi informacje nt. jego komputera. Wyszukuje zmapowane dyski sieciowe, podłączone drukarki, przeszukuje 
dyski w poszukiwaniu plików PST oraz wyświetla informacje o plikach PST podłączonych do profilu Outlooka. Dodatkowo skrypt zapisuje te informacje w plikach .dat 
(format CSV) i kopiuje je do wskazanego udziału sieciowego. Skrypt jest bezparametrowy i nie wymaga interakcji ze strony użytkownika.

.NOTES   
Name       : Get-ComputerInfo.ps1
Author     : Szymon Zdziechowiak
Version    : 1.00
DateCreated: 2019-11-18
DateUpdated: 2019-11-18
Email      : szymon.zdziechowiak@lumileds.com

.EXAMPLE
Get-ComputerInfo.ps1

Description: Wywoła działanie skryptu. Skrypt utworzy w c:\temp pliki .dat (CSV) z informacjami o podłączonych dyskach sieciowych, drukarkach, plikach PST oraz 
skopiuje je do udziału sieciowego $NetworkLogDir. Po zebraniu informacji skrypt otworzy w domyślnej przeglądarce stronę raportu z zebranymi informacjami.

#>

## Inicjacja zmiennych

$UsrName = $env:USERNAME
$CompName = $env:COMPUTERNAME
$CurrentDomain = $env:USERDOMAIN
$UserReportFileHTML = "c:\temp\" + $UsrName + "_UserEnvSummary.html"
$NetworkDrivesFile = "c:\temp\" + $UsrName + "_UserNetworkDrives.dat"
$NetworkPrintersFile = "c:\temp\" + $UsrName + "_UserNetworkPrinters.dat"
$NetworkPSTOutlookFile = "c:\temp\" + $UsrName + "_OutlookPstFiles.dat"
$NetworkPSTFile = "c:\temp\" + $UsrName + "_AllPstFiles.dat"
$NetworkLogDir = "\\plrpabser1vwfs1\data\30_PUBLIC\!_Wymiana_komputerów\" + $UsrName # Ścieżka sieciowa do zapisu pliku logu

### Sprawdzenie istnienia katalogu c:\temp\

if (!(Test-Path -Path "C:\temp" )) {
    New-Item -ItemType directory -Path "c:\temp"
    }

### Sprawdzenie istnienia katalogu indywidualnego w Issue Log

if (!(Test-Path -Path $NetworkLogDir )) {
    New-Item -ItemType directory -Path $NetworkLogDir
    }

### Odczyt informacji o zmapowanych dyskach

if (Test-Path ($NetworkDrivesFile)) {
    Remove-Item $NetworkDrivesFile
    }

$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('CurrentUser', $CompName)

$ReportDrives = @()

$NetKey = $Reg.OpenSubKey("Network")

    # Jeśli w gałęzi NETWORK są jakieś zmapowane dyski to wydobywamy te gałęzie
    if ($NetKey.SubKeyCount -gt 0) {
        $DriveKeys = $NetKey.GetSubKeyNames()
        for ($n = 0; $n -lt $DriveKeys.Length; $n++) {
            $DriveKey = $Reg.OpenSubKey("Network\\$($DriveKeys[$n])")

            # Utworzenie tabeli z listą dysków
            $hash = [ordered] @{
                MappedLocation = $DriveKey.GetValue("RemotePath")
                DriveLetter    = $DriveKeys[$n]
            }

            $objDriveInfo = New-Object PSObject -Property $hash

            $ReportDrives += $objDriveInfo

        }
    }
    
$ReportDrives | Export-Csv -NoType $NetworkDrivesFile

Write-Host "Working... 40%"

# kopiujemy plik .dat do sieci

Copy-Item $NetworkDrivesFile $NetworkLogDir -Force


### Odczyt i eksport danych o drukarkach sieciowych

if (Test-Path ($NetworkPrintersFile)) {
    Remove-Item $NetworkPrintersFile
}

$ReportPrinters = @()

$colPrinters = Get-WmiObject -Class Win32_Printer | Where-Object {$_.Network} | Select-Object Name
foreach ($objPrinter in $colPrinters) {
    
    $hashPrinter = [ordered] @{
        Location = $objPrinter.Name
    }

    $objPrinterInfo = New-Object PSObject -Property $hashPrinter

    $ReportPrinters += $objPrinterInfo
}

$ReportPrinters | Export-Csv -NoType $NetworkPrintersFile

Write-Host "Working... 50%"

# kopiujemy plik .dat do sieci

Copy-Item $NetworkPrintersFile $NetworkLogDir -Force


## Wyszukiwanie plików .pst w profilu Outlook

$ReportPSTOutlookFiles = @()

$Outlook = New-Object -comObject Outlook.Application 

$PSTOutlookFiles = $Outlook.Session.Stores | Where-Object { ($_.FilePath -like '*.PST') } | Select-Object FilePath 

ForEach ($objPSTOutlookFile in $PSTOutlookFiles) {
    $hashPSTOutlookFiles = [ordered] @{
        Katalog = $objPSTOutlookFile.FilePath
    }

    $objPSTInfo = New-Object PSObject -Property $hashPSTOutlookFiles

    $ReportPSTOutlookFiles += $objPSTInfo
}

$ReportPSTOutlookFiles | Export-Csv -NoType $NetworkPSTOutlookFile -Encoding UTF8

Write-Host "Working... 75%"

# kopiujemy plik .dat do sieci

Copy-Item $NetworkPSTOutlookFile $NetworkLogDir -Force


## Wyszukiwanie wszystkich plików .pst na dyskach twardych

$ReportPSTFiles = @()

$PSTFiles = Get-WMIObject Win32_LogicalDisk -filter "DriveType = 3" | Select-Object DeviceID | ForEach-Object {Get-Childitem ($_.DeviceID + "\") -include *.pst -recurse -ErrorAction SilentlyContinue | select-object Directory, Name}

ForEach ($objPSTFile in $PSTFiles) {
    $hashPSTFiles = [ordered] @{
        KatalogiPST = $objPSTFile.Directory
        PlikiPST    = $objPSTFile.Name
    }
    $objPSTFileInfo = New-Object PSObject -Property $hashPSTFiles
    $ReportPSTFiles += $objPSTFileInfo
}

$ReportPSTFiles | Export-Csv -NoType $NetworkPSTFile -Encoding UTF8

Write-Host "Working... 99%"

# kopiujemy plik .dat do sieci

Copy-Item $NetworkPSTFile $NetworkLogDir -Force

## Tworzenie zmiennych dla raportu HTML

$dyski = $ReportDrives | Select-Object @{Name = "Dyski"; Expression = {$_.DriveLetter}}, @{Name = "Foldery"; Expression = {$_.MappedLocation}}

$drukarki = $ReportPrinters | Select-Object @{Name = "Drukarki sieciowe"; Expression = {$_.Location}}

$plikipst = $ReportPSTFiles | Select-Object @{Name = "Katalog pliku PST"; Expression = {$_.KatalogiPST}}, @{Name = "Plik PST z archiwum poczty"; Expression = {$_.PlikiPST}}

$plikipstoutlook = $ReportPSTOutlookFiles | Select-Object @{Name = "Pliki PST w Twoim Outlooku"; Expression = {$_.Katalog}}

$title = "$CompName $UsrName"

$pre = "<div id=titleFont>Koniecznie wydrukuj tę stronę!</div><br><div id=normalText>Ta strona zawiera informacje o Twoich połączeniach sieciowych `
i miejscach przechowywania Twojej poczty e-mail.`
</div><br><div id=titleFont>Nazwa Twojego komputera: $($CompName)</div><br><div id=titleFont>Twoja nazwa użytkownika: $CurrentDomain\$UsrName</div><br>"

$post = "<br><div id=normalText>Strona wygenerowana $(Get-Date -UFormat "%Y-%m-%d %T")</div><br>"

# Generowanie HTML

$dyski | Convertto-html -Title $title -PreContent $pre -CssUri .\Get-ComputerInfo.css | Out-File $UserReportFileHTML -Encoding UTF8

$drukarki | ConvertTo-Html  -Property "Drukarki sieciowe" | Out-File $UserReportFileHTML -Append -Encoding UTF8

$plikipst | Convertto-Html | Out-File $UserReportFileHTML -Append -Encoding UTF8

$plikipstoutlook | Convertto-html -Property "Pliki PST w Twoim Outlooku" -PostContent $post | Out-File $UserReportFileHTML -Append -Encoding UTF8

Write-Host "DONE!"

Copy-Item $UserReportFileHTML $NetworkLogDir -Force
Copy-Item "C:\temp\Get-ComputerInfo.css" $NetworkLogDir

Invoke-Item $UserReportFileHTML

# Remove-Item -LiteralPath $MyInvocation.MyCommand.Path -Force