##############################################################
# 
# Считается общее кол-во КО, с разбивкой по категориям
# на основе списка MS Sharepoint 
# http://www.kaluga.cbr.ru/deprts/nadz/Lists/List/view1.aspx
# 
# 
# Заказчик: Отдел Банковского надзора
# Исполнитель: Астахов А.Б.
# Начато: 28.02.2011
# 
#
##############################################################
Function Get-HTMLHeader{
$dtToday = $(Get-Date).ToShortDateString()

$header = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
$header += '<html xmlns="http://www.w3.org/1999/xhtml">'
$header += '<head>'
$header += '<meta http-equiv="Content-Language" content="en-us" />'
$header += '<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />'

$header += '<style type="text/css">'
$header += '.bd1 {'
$header += 				'border-style: ridge solid groove solid;'
$header += 				'border-width: 1px;'
$header += 				'border-color: #CC33FF #CC33FF #C0C0C0 #CC33FF;'
$header += 				'background: #E6E6FF;'
$header += 				'padding: 5px;'
$header += 				'font-family: Tahoma;'
$header += 				'font-size: small;'
$header += 				'text-align: center;'
$header += 				'margin-top: auto;'
$header += 				'margin-bottom: auto;'
$header += 				'margin-left: auto;'
$header += '}'
$header += '.bd2 {'
$header +=  			'border-style: ridge solid groove solid;'
$header += 				'border-width: 1px;'
$header += 				'border-color: #CC33FF #CC33FF #C0C0C0 #CC33FF;'
$header += 				'background: #C3DAF9;'
$header += 				'padding: 5px;'
$header += 				'font-family: Tahoma; '
$header += 				'font-size: small;'
$header += 				'text-align: center;'
$header += 				'margin-top: auto;'
$header += 				'margin-bottom: auto;'
$header += 				'margin-left: auto;'
$header += '}'
$header += '.Cell1 {border: 1px solid #D6E3F1;}'
$header += '.rnav tr:hover {background: #FFFFCC;}'
$header += '.rnav1 tr:hover {background: #D5E4F2;}'
$header += '.tbli1 td, th+th {border-bottom: 1px #FFFF66 solid;}'
$header += '.z1 {background: #9FD5EB; border: 1px solid #FFF; padding: 5px; font-family:Tahoma;font-size:large; text-align:center}'
$header += '.z2 {background:#C3DAF9; border:1px solid #3B619C; padding:5px; font-family:Tahoma;font-size:small;text-align:center}'
$header += '.z3 {background:#154A93; border:1px solid #3B619C; padding:5px; font-family:Tahoma;font-size:small;text-align:center;color:#E3F4A8}'
$header += "</style>"


$header += "<title>Общая статистика КО на $dtToday</title>"
$header += "</head>"
$header += "<body>"

$header += '<table style="width: 100%" >'
$header += 				'<thead>'
$header += 				'<tr class="z1" >'
$header += 								'<td colspan="12" >'
$header += 								'Общая статистика КО '
$header += 								"на $dtToday</td>"
$header +=				'</tr>'
$header +=				'<tr class="z2">'
$header +=								'<td>Кредитные<br/>организации<br/>региона</td>'
$header +=								'<td >Филиалы</td>'
$header +=								'<td >Представительства</td>'
$header +=								'<td>Дополнительные<br/>офисы</td>'


$header +=								'<td>Операционные<br/>кассы вне<br/>кассового узла</td>'
$header +=								'<td>Операционные<br/>офисы</td>'

$header +=								'<td>Кредитно-<br/>Кассовые<br/>офисы</td>'
$header +=								'<td>Передвижные<br/>пункты<br/>кассовых операций</td>'
$header +=								'<td>Общее<br/>количество КО<br/>Калужского региона</td>'
$header +=								'<td>Структурные подразделения КО<br/>(филиалов) за пределами<br/>Калужской области</td>'
$header +=								'<td>Общее количество КО,<br/>в т.ч. за пределами региона<br/>(без учета КО с отозванной лицензией)</td>'
$header +=								'<td>КО с<br/>отозванной<br/>лицензией</td>'

								
$header +=				'</tr>'
#$header +=				'<tr class="z3">'
#$header +=								'<td>1</td>'
#$header +=								'<td>2</td>'
#$header +=								'<td>3</td>'
#$header +=								'<td>4</td>'
#$header +=								'<td>5</td>'
#$header +=								'<td>6</td>'
#$header +=								'<td>7</td>'
#$header +=								'<td>8</td>'
#$header +=								'<td>9</td>'
#$header +=								'<td>10</td>'
#$header +=								'<td>11</td>'
#$header +=								'<td>12</td>'
#$header +=								'<td>13</td>'
#$header +=								'<td>14</td>'
#$header +=								'<td>15</td>'



#$header +=				'</tr>'
$header +=				'</thead>'
$header +=				'<tbody valign="top" class="rnav">'


Return $header
}
#-------------------------------------------------------------------------------------
Function Get-HTMLFooter{
$footer =				"</tbody>"
				

$footer +=			"</table>"

$footer +=		"</body>"

$footer +="</html>"
Return $footer
}
#-------------------------------------------------------------------------------------
Function Get-HTMLReportBody ($KOCount, $FilCount, $Predstav, $DopOffice, $OperKassyVneKassUzla, $OperOffice, $KKOffice, $TotalCount,  $KOZaPredelami, $AllCount, $LicenseRevokedCount, $PPKO){

							# $BnkLogo $BnkLogoAlt $numPP  $NaimPiK  $NaimGolov  $OsnOtzLic  $RukTbl   $TelTbl  $UAddress  $PAddress  $Okpo  $Inn  $Kpp  $Ogrn  $OgrnDate  $KS4et  $Bik  $Okved
$tcell = '<td class="z3">'
$body = 				'<tr class="bd1">'
$body += $tcell+"$KOCount</td>"

$body += $tcell+"$FilCount</td>"
$body += $tcell+"$Predstav</td>"

$body += $tcell+"$DopOffice</td>"

$body += $tcell+"$OperKassyVneKassUzla</td>"
$body += $tcell+"$OperOffice</td>"
$body += $tcell+"$KKOffice</td>"
$body += $tcell+"$PPKO</td>"
$body += $tcell+"$TotalCount</td>"
$body += $tcell+"$KOZaPredelami</td>"
$body += $tcell+"$AllCount</td>"
$body += $tcell+"$LicenseRevokedCount</td>"

															
$body += 				'</tr>'

Return $body
}
#-------------------------------------------------------------------------------------
Start-Transcript ‘f:\AdminDir\SpisokKO\07C_TotalCount.log’ -force
$FileName = "f:\AdminDir\SpisokKO\"+$(Get-Date).Year+"."

if ($(Get-Date).Month -lt 10){
	$FileName += "0"
	} 
$FileName += [String]($(Get-Date).Month)+"."

if ($(Get-Date).Day -lt 10){
	$FileName += "0"
	} 
$FileName += [String]$(Get-Date).Day+"-07C_Общее_количество" 
# $FileName
$AlertEmailStr	= "dl.nadz.alert@kaluga.cbr.ru"
$ListLinkStr	= "http://www.kaluga.cbr.ru/deprts/nadz/Lists/List/DispForm.aspx?ID="
$ListLinkStr1	= "&Source=%2Fdeprts%2Fnadz%2Fdefault%2Easpx"
$AlertBankName	= @()
$AlertBankLinks	= @()
$AlertDateDovOver	= @()
$bankArr = @()
$rsymb  = [char][int]1

$KOZaPredelami			= 0

$TotalCount				= 0
$LicenseRevokedCount	= 0
$KOCount				= 0
$FilCount				= 0
$Predstav				= 0
$DopOffice				= 0
$OperKassyVneKassUzla	= 0
$OperOffice				= 0
$KKOffice				= 0
$PPKO                   = 0

$MyReport = Get-HTMLHeader

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
		
		
		for ($jj=0; $jj -lt $iCnt; $jj++){
			
			$spcurItem = $spList.Items.item($jj);
			# Калужские банки
			
			
			$licenseRevoked = ($spcurItem["Лицензия"] -eq "Отозвана")

			if (!$licenseRevoked) {

				if (!($spcurItem["Регион"] -eq "КО за пределами Калужской области")){
			
					if (($spcurItem["Регион"] -eq "Калужские банки") -and ($spcurItem["Головной банк"] -eq "Да")){
						$KOCount++
					}

					if ( ($spcurItem["Тип кредитной организации"] -eq "Банки") -and ($spcurItem["Регион"] -eq "Филиалы")){
						$FilCount++
					}

					if ( ($spcurItem["Тип кредитной организации"] -eq "Представительства")){
						$Predstav++
					}

					if ( ($spcurItem["Тип кредитной организации"] -eq "Дополнительные офисы")){
						$DopOffice++
					}

					if ( ($spcurItem["Тип кредитной организации"] -eq "Операционные кассы")){
						$OperKassyVneKassUzla++
						
					}

					if ( ($spcurItem["Тип кредитной организации"] -eq "Операционные офисы")){
						$OperOffice++
					}

					if ( ($spcurItem["Тип кредитной организации"] -eq "Кредитно-Кассовые офисы")){
						$KKOffice++
					}


					if ( ($spcurItem["Тип кредитной организации"] -eq "Передвижные пункты кассовых операций")){
						$PPKO++
					}


					$TotalCount++
				}
				else
				{
					$KOZaPredelami++
				}
			}else
			{
				$LicenseRevokedCount++
			
			}
			#read-host

		}

$AllCount  = $TotalCount+$KOZaPredelami
$MyReport += Get-HTMLReportBody $KOCount $FilCount $Predstav $DopOffice $OperKassyVneKassUzla $OperOffice $KKOffice $TotalCount  $KOZaPredelami $AllCount $LicenseRevokedCount $PPKO
$MyReport += Get-HTMLFooter
$Filename += "($TotalCount).html"
$MyReport | out-file -encoding UTF8 -filepath $Filename



Write-output "Program END"
Stop-Transcript