$profilename= $MyInvocation.MyCommand.Name

$ver = $host | select version

if ($ver.Version.Major -gt 1) {$host.Runspace.ThreadOptions = "ReuseThread"}

Add-PsSNapIn Microsoft.SharePoint.PowerShell

if ($host.Name -eq "ConsoleHost")
{

	$width = 80
	$sizeWindow = new-Object System.Management.Automation.Host.Size $width, 40
        $sizeBufer  = new-Object System.Management.Automation.Host.Size $width, 9999
	<#
	The buffer width can't be resized to be narrowed than the windows's current width
	plus the window's width can't be resized to be wider than the buffer current width
	#>
        $S = $Host.UI.RawUI
        if ($S.WindowSize.width -gt $width)
        {
		$s.WindowSize = $sizeWindow
		$s.BufferSize = $sizeBufer
		

	}
	else
	{
		
		$s.BufferSize = $sizeBufer
		$s.WindowSize = $sizeWindow
	}

        $s.ForegroundColor = "Yellow";
        $s.BackgroundColor = "DarkBlue"

        $s.windowTitle = "$env:computername"
}
