function LoadSnapin
{
	Add-PsSNapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

}

LoadSnapin

$nsite="http://s29sps.region.cbr.ru"
$SpSite = Get-SPSite $nsite

