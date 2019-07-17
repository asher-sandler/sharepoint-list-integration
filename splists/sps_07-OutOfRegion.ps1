##############################################################
# 
# Список КО формируется в виде HTML страницы
# на основе списка MS Sharepoint 
# http://www.kaluga.cbr.ru/deprts/nadz/Lists/List/view1.aspx
# c фильтром по полям
# ([Тип кредитной организации]=="Банки" и 
# [Регион]=="Филиалы") или ([Регион] == "Калужские банки")
# т.е. выбираем Калужские банки и Филиалы торонних Банков
# 
# 
# Заказчик: Отдел Банковского надзора
# Исполнитель: Астахов А.Б.
# Начато: 18.05.2010
# 
# Скрипт также проверяет если поле "Срок действия доверенности" или 
# "Срок действия доверенности 2"
# меньше или равно сегодняшняя дата +21 (предупреждение за 21 день)
# то посылается сообщение по электронной почте о том что истекло или 
# приближается истечение срока действия доверенности
# Адрес рассылки находится в переменной $AlertEmailStr
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


$header += "<title>Список действующих кредитных организаций и филиалов действующих КО на $dtToday</title>"
$header += "</head>"
$header += "<body>"

$header += '<table style="width: 100%" >'
$header += 				'<thead>'
$header += 				'<tr class="z1" >'
$header += 								'<td colspan="8" >'
$header += 								'Список Структурных подразделений КО (филиалов) за пределами Калужской области  '
$header += 								'по '
$header += 								"состоянию на $dtToday</td>"
$header +=				'</tr>'
$header +=				'<tr class="z2">'
$header +=								'<td>№ <br/>п/п</td>'
$header +=								'<td >Наименование <br/>(полное и краткое)</td>'
$header +=								'<td >Наименование головной кредитной организации, адрес</td>'
#-------------------------------------------------------------------------------------


#					$header +=								'<td>Вид лицензии /<br/>учредительные документы</td>'
#-------------------------------------------------------------------------------------
$header +=								'<td>Порядковый<br/>номер,<br/>присвоенный<br/>Банком России</td>'
$header +=								'<td>Дата<br/>открытия</td>'
#-------------------------------------------------------------------------------------
#					$header +=								'<td>№ фил.</td>'
#-------------------------------------------------------------------------------------
$header +=								'<td>Руководители</td>'
#-------------------------------------------------------------------------------------
#					$header +=								'<td>№ и дата<br/>доверенности</td>'
#					$header +=								'<td>Срок<br/>действия<br/>доверенности</td>'
#					$header +=								'<td>Телефон (Факс)</td>'
#					$header +=								'<td>Юридический<br/>адрес</td>'
#-------------------------------------------------------------------------------------
$header +=								'<td>Фактический<br/>адрес</td>'
$header +=								'<td>Телефон (Факс)</td>'
#-------------------------------------------------------------------------------------
#					$header +=								'<td>ОКПО</td>'
#					$header +=								'<td>ИНН/&nbsp;&nbsp;&nbsp;КПП</td>'
#					$header +=								'<td>ОГРН /<br/>серия, №,<br/>дата свидетельства</td>'
#					$header +=								'<td>к/счет (субсчет),<br/>наименование РКЦ</td>'
#					$header +=								'<td>БИК</td>'
#					$header +=								'<td>ОКВЭД</td>'
								
$header +=				'</tr>'
$header +=				'<tr class="z3">'
$header +=								'<td>1</td>'
$header +=								'<td>2</td>'
$header +=								'<td>3</td>'
$header +=								'<td>4</td>'
$header +=								'<td>5</td>'
$header +=								'<td>6</td>'
$header +=								'<td>7</td>'
$header +=								'<td>8</td>'


$header +=				'</tr>'
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
Function Get-HTMLRukovCells ($Rank, $Name){

$cells = 										'<tr class="tbli1">'
$cells+=														"<td>$Rank</td>"
$cells+=														"<td>$Name</td>"
$cells+=										'</tr>'

Return $cells
}
#-------------------------------------------------------------------------------------
Function Get-HTMLPhoneCells ($PhoneNumb){

$outStr='<table style="width: 100%" class="tbli1"><tbody class="rnav1">'

$PhoneNumb = Del-HTMLMarkup $PhoneNumb 


$wPhone = $PhoneNumb.Split(",")

$aPhone=@()
$PhoneStr = ""

# Убираем пустые
foreach ($q in $wPhone) {
		if ($($q.trim()).length -gt 0){
			$aPhone += $q
		
		}
}

# заполняем телефоны


for  ($kk=0; $kk -lt $aPhone.length;$kk++) {
         
        $PhoneStr += '<tr class="tbli1"><td>'+$aPhone[$kk]+'</td></tr>'
		
        }

$outStr += $PhoneStr+"</tbody></table>"

Return $outStr
}
#-------------------------------------------------------------------------------------
Function Get-RukovTable($rukov,$GlBuh){

$outStr='<table style="width: 100%" class="tbli1"><tbody class="rnav1">'

$Ruk = $rukov -replace("<br>","\")
$GLB = $GlBuh -replace("<br>","\")

$wRuk = $Ruk.Split("\")
$wGlb = $Glb.Split("\")

$aRuk=@()
$aGlb=@()
$glbstr = $upravStr = ""

# Убираем пустые
foreach ($q in $wRuk) {
		if ($($q.trim()).length -gt 0){
			$aRuk += $q
		
		}
}
foreach ($q in $wGlb) {
		if ($($q.trim()).length -gt 0){
			$aGlb += $q
		
		}
}
# заполняем управляющих


for  ($kk=0; $kk -lt $aRuk.length;$kk++) {
        # 
        $upravstr += '<tr class="tbli1"><td>'+$aRuk[$kk]+'</td></tr>'
		
        }



# заполняем бухгалтеров        


for  ($kk=0; $kk -lt $aGlb.length;$kk++) {
       
		$glbstr  += '<tr class="tbli1"><td>'+$aGlb[$kk]+'</td></tr>'
	
		
        } 

        

#write-host $upravstr
#write-host $glbstr
              
$outStr += $upravstr+$glbstr+"</tbody></table>"
# $outStr += "</tbody></table>"
Return $outStr
}

#-------------------------------------------------------------------------------------
Function Get-HTMLReportBody ($BnkLogo, $BnkLogoAlt, $numbPP, $NaimPiK, $NaimGolov, $VidLic, $DatReg, $RegNom, $NomFil, $RukTbl, $NiD_Dov, $Srok_Dov, $Srok_Dov2, $TelTbl, $UAddress, $PAddress, $Okpo, $Inn, $Kpp, $Ogrn, $OgrnDate, $KS4et, $Bik, $Okved){
$BnkLogoAlt = $BnkLogoAlt.replace('"','&quot;')
$tcell = '<td class="Cell1">'
if (($numbPP % 2) -eq 1){
	$body = 				'<tr class="bd1">'
	}
else{
	$body = 				'<tr class="bd2">'
	}

$body += $tcell+"$numbPP</td>"
$body += $tcell+"$NaimPiK<br/><br/><img src="+'"'+$BnkLogo+'" alt="'+$BnkLogoAlt+'" /></td>'
$body += $tcell+"$NaimGolov</td>"
#			$body += $tcell+"$VidLic</td>"
$body += $tcell+"$RegNom</td>"
$body += $tcell+"$DatReg</td>"
#			$body += $tcell+"$NomFil</td>"
$body += $tcell+"$RukTbl</td>"
#			$body += $tcell+"$NiD_Dov</td>"
#			$body += $tcell+"$Srok_Dov<br/>$Srok_Dov2</td>"
#			$body += $tcell+"$UAddress</td>"
$body += $tcell+"$PAddress</td>"
$body += $tcell+"$TelTbl</td>"
#			$body += $tcell+"$Okpo</td>"
#			$body += $tcell+"$Inn<br/>/<br/>$Kpp</td>"
#			$body += $tcell+"$Ogrn<br/><br/>$OgrnDate</td>"
#			$body += $tcell+"$KS4et</td>"
#			$body += $tcell+"$Bik</td>"
#			$body += $tcell+"$Okved</td>"	
															
$body += 				'</tr>'

Return $body
}
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
Function Get-DateDovOver ($SrokDov){


$AlertDate = $SrokDov.AddDays(-21)
$NowDate   = Get-Date
$IsAlert   = ($AlertDate -le $NowDate)

Return $IsAlert
}
#-------------------------------------------------------------------------------------
Function Get-HTMLMailBody ($BankName, $BankLinks,	$DateDov){


	$HTMLMailBody	 = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
	$HTMLMailBody	+= '<html xmlns="http://www.w3.org/1999/xhtml">'
	$HTMLMailBody	+= '<head>'
	$HTMLMailBody	+= '<meta http-equiv="Content-Language" content="ru" />'
	$HTMLMailBody	+= '<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />'
	$HTMLMailBody	+= '<title>Список филиалов с истекшим сроком действия доверенности</title>'
	$HTMLMailBody	+= '</head>'
	$HTMLMailBody	+= '<body>'
	$HTMLMailBody	+= '<table style="width: 100%">'
	$HTMLMailBody	+= '<tr><td colspan="3" style="font-family: Arial, Helvetica, sans-serif; text-align: center; font-weight: 700; border-left-style: solid; border-left-color: #C0C0C0; border-right-style: solid; border-top-style: solid; border-top-color: #C0C0C0; border-bottom-style: solid; background-color: #C0C0C0;">Список филиалов с истекающим сроком действия доверенности. '
	$HTMLMailBody	+= "Отчет за "+$(Get-Date).ToShortDateString()
	$HTMLMailBody	+= '</td></tr>'
	$HTMLMailBody	+= '<tr style="text-align: center; font-family: Arial, Helvetica, sans-serif; font-size: small; border-style: solid; border-color: #FFFFCC; background-color: #FFFFCC"><td>№<br/>п/п</td><td>Наименование</td><td>Срок<br/>действия<br/>доверенности</td></tr>'

    for ($jj=0;$jj -lt $BankName.length;$jj++){
		$HTMLMailBody	+= '<tr style="font-family: Arial, Helvetica, sans-serif; font-size: small; background-color: #C9DDFC"><td style="text-align: center">'
		$HTMLMailBody	+= [string]$($jj+1)
		$HTMLMailBody	+= '</td><td><a href="'
		$HTMLMailBody	+= $BankLinks[$jj]
		$HTMLMailBody	+= '">'
		$HTMLMailBody	+= $BankName[$jj]
		$HTMLMailBody	+= '</a></td><td>'
		$HTMLMailBody	+= $DateDov[$jj]
		$HTMLMailBody	+= '</td></tr>'

		
		}

	$HTMLMailBody	+= '</table></body></html>'

Return $HTMLMailBody

}
#-------------------------------------------------------------------------------------
Start-Transcript ‘f:\AdminDir\SpisokKO\07OutOfRegion.log’ -force
$FileName = "f:\AdminDir\SpisokKO\"+$(Get-Date).Year+"."

if ($(Get-Date).Month -lt 10){
	$FileName += "0"
	} 
$FileName += [String]($(Get-Date).Month)+"."

if ($(Get-Date).Day -lt 10){
	$FileName += "0"
	} 
$FileName += [String]$(Get-Date).Day+"-07_Список_Структурных_подразделений_КО_(филиалов)_за_пределами_Калужской_области" 
# $FileName
$AlertEmailStr	= "dl.nadz.alert@kaluga.cbr.ru"
$ListLinkStr	= "http://www.kaluga.cbr.ru/deprts/nadz/Lists/List/DispForm.aspx?ID="
$ListLinkStr1	= "&Source=%2Fdeprts%2Fnadz%2Fdefault%2Easpx"
$AlertBankName	= @()
$AlertBankLinks	= @()
$AlertDateDovOver	= @()

$bankArr = @()

$rsymb  = [char][int]1
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
		$icnt1=$icnt2=0;
		
		$numPP=1
		
		
		# КО за пределами Калужской области
		for ($jj=0; $jj -lt $iCnt; $jj++){
			
			$spcurItem = $spList.Items.item($jj);
			# Калужские банки
			
			
			
			
			#$spcurItem["Банк"]
			#$spcurItem["Регион"]
			#$spcurItem["Тип кредитной организации"]
			
			
			
			$HTMLrowString = ""
			
			$BnkLogoAlt = $BnkLogo = $NaimPiK = $NaimGolov = $VidLic = $DatReg = $RegNom = $NomFil = $RukTbl = $NiD_Dov = $Srok_Dov = $Srok_Dov2= $TelTbl = $UAddress = $PAddress = $Okpo = $Inn = $Kpp = $Ogrn = $OgrnDate = $KS4et = $Bik = $Okved = ""
			
			$bnkID		= 0
			$bnkID		= $spcurItem["ИД"]
			$BnkLogo 	= $spcurItem["БанкЛого"]
			$NaimPiK	= Del-HTMLMarkup $spcurItem["Наименование (полное и краткое)"]
			$NaimGolov	= Del-HTMLMarkup $spcurItem["Наименование головной кредитной организации, адрес"]
			$DatReg		= $($spcurItem["Дата регистрации"]).ToShortDateString()
			$RegNom		= $spcurItem["Порядковый номер, присвоенный Банком России"]
			$NomFil		= $spcurItem["№ фил"]
			$RukTbl		= Get-RukovTable $spcurItem["Руководитель"] $spcurItem["Гл. бухгалтер"]
			$NiD_Dov	= Del-HTMLMarkup $spcurItem["№ и дата доверенности"]
			$Srok_Dov	= [string]($spcurItem["Срок действия доверенности"])
			if ($Srok_Dov.length -gt 0){
			    $Srok_Dov = $([DateTime]($Srok_Dov)).ToShortDateString()
			}
			$Srok_Dov2  = [string]($spcurItem["Срок действия доверенности 2"])
			if ($Srok_Dov2.length -gt 0){
			    $Srok_Dov2 = $([DateTime]($Srok_Dov2)).ToShortDateString()
			}
			if ($spcurItem["Телефон"].length -gt 0){
				$TelTbl		= Get-HTMLPhoneCells $spcurItem["Телефон"]
				}
			else{
				$TelTbl = ""
				}	
			$UAddress	= Del-HTMLMarkup $spcurItem["Юридический адрес"]
			$PAddress	= Del-HTMLMarkup $spcurItem["Фактический адрес"]
			$Okpo		= $spcurItem["ОКПО"]
			$Inn		= $spcurItem["ИНН"] 
			$Kpp		= $spcurItem["КПП"]
			$Ogrn		= $spcurItem["ОГРН"]
			$OgrnDate	= Del-HTMLMarkup $spcurItem["серия, №, дата свидетельства"]
			$KS4et		= Del-HTMLMarkup $spcurItem["к/счет (субсчет), наименование РКЦ"]
			$Bik		= $spcurItem["БИК"]
			$Okved		= $spcurItem["ОКВЭД"]
			 
			
			$licenseRevoked = ($spcurItem["Лицензия"] -eq "Отозвана")

			if (!$licenseRevoked) {
			
				
				if (($spcurItem["Регион"] -eq "КО за пределами Калужской области")){
						# проверяем срок доверенности
					
						$VidLic = "Положение"
						$BnkLogoAlt = $BnkLogo.Split(",")[1].Trim()
						$BnkLogo = $BnkLogo.Split(",")[0].Trim()
				
				
						$bankArr +=  $spcurItem["Банк"] + $spcurItem["Заполните для сортировки"] + $rsymb+$BnkLogo +$rsymb+ $BnkLogoAlt+$rsymb+  $NaimPiK +$rsymb+ $NaimGolov+$rsymb+  $VidLic +$rsymb+ $DatReg+$rsymb+  $RegNom+$rsymb+  $NomFil+$rsymb+  $RukTbl+$rsymb+  $NiD_Dov +$rsymb+ $Srok_Dov +$rsymb+ $Srok_Dov2 +$rsymb+$TelTbl+$rsymb+  $UAddress +$rsymb+ $PAddress+$rsymb+  $Okpo+$rsymb+  $Inn+$rsymb+  $Kpp+$rsymb+  $Ogrn+$rsymb+  $OgrnDate+$rsymb+  $KS4et+$rsymb+  $Bik+$rsymb+  $Okved
						# $MyReport += Get-HTMLReportBody $BnkLogo $BnkLogoAlt $numPP  $NaimPiK  $NaimGolov  $VidLic  $DatReg  $RegNom  $NomFil  $RukTbl  $NiD_Dov  $Srok_Dov  $Srok_Dov2 $TelTbl  $UAddress  $PAddress  $Okpo  $Inn  $Kpp  $Ogrn  $OgrnDate  $KS4et  $Bik  $Okved
						$icnt2++
						# $numPP++
				}
			}
			
			#read-host

		}		
$bsort= $bankArr | sort


$numPP=1

for ($i=0;$i -lt $bsort.Count;$i++){

	
	$strArr = $bsort[$i].split($rsymb)
    $BnkLogo 	= $strArr[1]
	$BnkLogoAlt	= $strArr[2]
	$NaimPiK	= $strArr[3]
	$NaimGolov	= $strArr[4]
	$VidLic		= $strArr[5]
	$DatReg		= $strArr[6]
	$RegNom		= $strArr[7]
	$NomFil		= $strArr[8]
	$RukTbl		= $strArr[9]
	$NiD_Dov	= $strArr[10]
	$Srok_Dov	= $strArr[11]
	$Srok_Dov2	= $strArr[12]
	$TelTbl		= $strArr[13]
	$UAddress	= $strArr[14]
	$PAddress	= $strArr[15]
	$Okpo		= $strArr[16]
	$Inn		= $strArr[17]
	$Kpp		= $strArr[18]
	$Ogrn		= $strArr[19]
	$OgrnDate	= $strArr[20]
	$KS4et		= $strArr[21]
	$Bik		= $strArr[22]
	$Okved		= $strArr[23]
	$MyReport += Get-HTMLReportBody $BnkLogo $BnkLogoAlt $numPP  $NaimPiK  $NaimGolov  $VidLic  $DatReg  $RegNom  $NomFil  $RukTbl  $NiD_Dov  $Srok_Dov  $Srok_Dov2 $TelTbl  $UAddress  $PAddress  $Okpo  $Inn  $Kpp  $Ogrn  $OgrnDate  $KS4et  $Bik  $Okved


	$numPP++
	
}
		

$MyReport+=Get-HTMLFooter
$Filename+="($icnt2).html"
$MyReport | out-file -encoding UTF8 -filepath $Filename


Write-output "Program END"
Stop-Transcript