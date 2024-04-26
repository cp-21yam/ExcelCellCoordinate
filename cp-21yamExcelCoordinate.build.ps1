task Test {
    Import-Module Pester
    Invoke-Pester .\cp-21yamExcelCoordinate.Tests.ps1

    Import-Module PSScriptAnalyzer
    Invoke-ScriptAnalyzer -Path .\cp-21yamExcelCoordinate.psm1
}
task Release {
    Publish-PSResource -Path .\cp-21yamExcelCoordinate.psd1 -Repository PSGallery -ApiKey 'oy2ijintx2menomqvzzqub2jut3dex2endgt2pvgrrfopq'
}
