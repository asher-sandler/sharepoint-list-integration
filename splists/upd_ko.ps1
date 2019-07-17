#-------------------------------------------------------------------------------------
Function Del-HTMLMarkup	($htmlStr){
$outstr=""
$ISinclude = $true
for ($kk=0;$kk -lt $htmlStr.length; $kk++){
	
	if ($htmlStr.substring($kk,1)  -eq "<"){
	    $ISinclude=$false
		}
	if ($htmlStr.substring($kk,1)  -eq ">"){
	    $ISinclude=$true
	    continue
		}
	
    if ($ISinclude){
		$outstr += $htmlStr.substring($kk,1)
		}
    }  # end-for
Return $outstr     
}
#-------------------------------------------------------------------------------------


		$env:SPpath = "${env:CommonProgramFiles}\Microsoft Shared\web server extensions\12\"
		[System.Reflection.Assembly]::LoadFrom("$env:SPPath\ISAPI\Microsoft.SharePoint.dll")
        write-host open web
		# открываем web
		$nsite="http://www.kaluga.cbr.ru/deprts/nadz/"
		$SpSite = New-Object -TypeName "Microsoft.SharePoint.SPSite" -ArgumentList $nsite;
		$spweb=$spsite.OpenWeb();
        write-host open Sharepoint list
		# открываем  список
		$nlist="http://www.kaluga.cbr.ru/deprts/nadz/Lists/List/view1.aspx"
		$splist=$spweb.getlist($nlist);
		$iCnt = $splist.Items.Count;
		# $icnt;
		$icnt1=$icnt2=0;
		
		$numPP=1
		
		
		for ($jj=0; $jj -lt $iCnt; $jj++){
			
			$spcurItem = $spList.Items.item($jj);
			# Калужские банки
			
			$address = $spcurItem["Юридический адрес"]
			$bank = $spcurItem["Банк"]
			
			if (!($spcurItem["Регион"] -eq "Калужские банки") -and ($spcurItem["Головной банк"] -eq "Да") -or (($spcurItem["Тип кредитной организации"] -eq "Банки") -and ($spcurItem["Регион"] -eq "Филиалы"))){
				
				$ua =  Del-HTMLMarkup $spcurItem["Юридический адрес"]
				$fa =  Del-HTMLMarkup $spcurItem["Фактический адрес"]
				$id = $spcurItem["ИД"]
				
				$ua =  $ua.trim()
				$fa =  $fa.trim()
				
				if ( (($ua.length -gt 0) -and ($fa.length -eq 0)) ){
					
				    write-host $spcurItem["Наименование (полное и краткое)"],$id 
					write-host 'Юридический адрес = ', $ua
					write-host 'Фактический адрес = ', $fa
					
					
					$spcurItem["Юридический адрес"] = ""
					$spcurItem["Фактический адрес"] = $ua
					$spcurItem.Update();
					
					
					
					}
				
				
			}
			
			#$spcurItem["Банк"]
			#$spcurItem["Регион"]
			#$spcurItem["Тип кредитной организации"]
			
			
			#read-host

		}

