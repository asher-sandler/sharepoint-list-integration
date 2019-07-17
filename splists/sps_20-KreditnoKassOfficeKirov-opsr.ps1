﻿##############################################################
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


$header += "<title>Список действующих Кредитно-Кассовых офисов кредитных организаций Кировского региона по состоянию на $dtToday</title>"
$header += "</head>"
$header += "<body>"

$header += '<table style="width: 100%" >'
$header += 				'<thead>'
$header += 				'<tr class="z1" >'
$header += 								'<td colspan="4" >'
$header += 								'Список действующих Кредитно-Кассовых офисов '
$header += 								'кредитных организаций Кировского региона по '
$header += 								"состоянию на $dtToday</td>"
$header +=				'</tr>'
$header +=				'<tr class="z2">'
$header +=								'<td>№ <br/>п/п</td>'
$header +=								'<td >Наименование <br/>(полное и краткое)</td>'
#-------------------------------------------------------------------------------------
#					$header +=								'<td >Наименование головной кредитной организации, адрес</td>'
#					$header +=								'<td>Вид лицензии /<br/>учредительные документы</td>'
#-------------------------------------------------------------------------------------
$header +=								'<td>Дата<br/>открытия</td>'
#-------------------------------------------------------------------------------------
#					$header +=								'<td>Рег.<br/>№</td>'
#					$header +=								'<td>№ фил.</td>'
#-------------------------------------------------------------------------------------
#					$header +=								'<td>Руководители</td>'
#-------------------------------------------------------------------------------------
#					$header +=								'<td>№ и дата<br/>доверенности</td>'
#					$header +=								'<td>Срок<br/>действия<br/>доверенности</td>'
#					$header +=								'<td>Телефон (Факс)</td>'
#					$header +=								'<td>Юридический<br/>адрес</td>'
#-------------------------------------------------------------------------------------
$header +=								'<td>Адрес</td>'
#					$header +=								'<td>Телефон (Факс)</td>'
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
$body += $tcell+"$NaimPiK</td>"
#			$body += $tcell+"$NaimPiK<br/><br/><img src="+'"'+$BnkLogo+'" alt="'+$BnkLogoAlt+'" /></td>'
#			$body += $tcell+"$NaimGolov</td>"
#			$body += $tcell+"$VidLic</td>"
$body += $tcell+"$DatReg</td>"
#			$body += $tcell+"$RegNom</td>"
#			$body += $tcell+"$NomFil</td>"
#			$body += $tcell+"$RukTbl</td>"
#			$body += $tcell+"$NiD_Dov</td>"
#			$body += $tcell+"$Srok_Dov<br/>$Srok_Dov2</td>"
#			$body += $tcell+"$UAddress</td>"
$body += $tcell+"$PAddress</td>"
#			$body += $tcell+"$TelTbl</td>"
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
Function Get-HTMLReportBodyShort ($numbPP,$str){
$tcell = '<td class="Cell1">'
if (($numbPP % 2) -eq 1){
	$body = 				'<tr class="bd1">'
	}
else{
	$body = 				'<tr class="bd2">'
	}
$body += $tcell+"$numbPP</td>"
$body += $tcell+$($str.split(";"))[3]+"</td>"
$body += $tcell+$($str.split(";"))[4]+"</td>"
$body += $tcell+$($str.split(";"))[5]+"</td>"

$body += 				'</tr>'

Return $body
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
Function Filt-Address ($address,$a_addr, $a_region)
{
		$found = ""

		#write-host $address
		#write-host $a_addr
		#read-host


		for ($ii = 0; $ii -lt $a_addr.Count-1; $ii++)
		{
		   if ( $address.Contains($a_addr[$ii])){

		       $found = $a_region[$ii]
			   break;
		   }
		}
    Return $found
}
#-------------------------------------------------------------------------------------
Start-Transcript ‘f:\AdminDir\SpisokKO\20KKofficeKirov_Opsr.log’ -force
$FileName = "f:\AdminDir\SpisokKO\"+$(Get-Date).Year+"."


if ($(Get-Date).Month -lt 10){
	$FileName += "0"
	} 
$FileName += [String]($(Get-Date).Month)+"."

if ($(Get-Date).Day -lt 10){
	$FileName += "0"
	} 
$FileName += [String]$(Get-Date).Day+"-20_Кред_касс_офис_Киров_ОПСиР" 
# $FileName

$region_spr = GET-CONTENT f:\AdminDir\SpisokKO\spavochniki\kirov_sprav.txt
$region_str  = ""

foreach ($spr_item in $region_spr)
{     
      $a1 = $spr_item.Split("|")[0]
      $addr_filter += $a1 +"|"
	  $a2 = $spr_item.Split("|")[1]
	  $region_str  += $a2 +"|"

	  
}
$addr = @()
$region  = @()

$addr =  $addr_filter.Split("|")
$region  = $region_str.Split("|")

$AlertEmailStr	= "dl.nadz.alert@kaluga.cbr.ru"
$ListLinkStr	= "http://www.kaluga.cbr.ru/deprts/nadz/Lists/List/DispForm.aspx?ID="
$ListLinkStr1	= "&Source=%2Fdeprts%2Fnadz%2Fdefault%2Easpx"
$bankArr = @()
$filArr = @()
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
		
		
		
		
		
		for ($jj=0; $jj -lt $iCnt; $jj++){
			
			$spcurItem = $spList.Items.item($jj);
			# Калужские банки
			
			
			
			
			#$spcurItem["Банк"]
			#$spcurItem["Регион"]
			#$spcurItem["Тип кредитной организации"]

			$reginAddress = Filt-Address $spcurItem["Фактический адрес"] $addr $region

			if ($reginAddress.Length -gt 0){   # фактический адрес, который нам нужен
			    write-host $reginAddress
				$licenseRevoked = ($spcurItem["Лицензия"] -eq "Отозвана")

				if (!$licenseRevoked) {
			
					$HTMLrowString = ""
			
					$BnkLogoAlt = $BnkLogo = $NaimPiK = $NaimGolov = $VidLic = $DatReg = $RegNom = $NomFil = $RukTbl = $NiD_Dov = $Srok_Dov = $Srok_Dov2= $TelTbl = $UAddress = $PAddress = $Okpo = $Inn = $Kpp = $Ogrn = $OgrnDate = $KS4et = $Bik = $Okved = ""
			
					$bnkID		= 0
					$bnkID		= $spcurItem["ИД"]
					$BnkLogo 	= $spcurItem["БанкЛого"]
					$NaimPiK	= Del-HTMLMarkup $spcurItem["Наименование (полное и краткое)"]
					$NaimGolov	= Del-HTMLMarkup $spcurItem["Наименование головной кредитной организации, адрес"]
					$DatReg		= $($spcurItem["Дата регистрации"]).ToShortDateString()
					$RegNom		= $spcurItem["Рег.№"]
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
			 
					if ( ($spcurItem["Тип кредитной организации"] -eq "Кредитно-Кассовые офисы")){
				
						if (!($spcurItem["Регион"] -eq "КО за пределами Калужской области")){
							$VidLic = "рублевая и валютная /<br />Устав"
                                                
				
					   
					
				
							#'$numPP		= '+$numPP
							$BnkLogoAlt = $BnkLogo.Split(",")[1].Trim()
							$BnkLogo = $BnkLogo.Split(",")[0].Trim()
				
				
							$NaimPiK = $NaimPiK.Replace(";"," ")
							$PAddress = $PAddress.Replace(";"," ")
					
							$bankArr += $reginAddress+";"+$spcurItem["Банк"] + ";"+$spcurItem["Заполните для сортировки"] + ";"+$NaimPiK+";"+$DatReg+";"+$PAddress
				
							#$MyReport += Get-HTMLReportBody $BnkLogo $BnkLogoAlt $numPP  $NaimPiK  $NaimGolov  $VidLic  $DatReg  $RegNom  $NomFil  $RukTbl  $NiD_Dov  $Srok_Dov $Srok_Dov2 $TelTbl  $PAddress  $UAddress  $Okpo  $Inn  $Kpp  $Ogrn  $OgrnDate  $KS4et  $Bik  $Okved
							$icnt1++
							#$numPP++
						}
				
					}
				}
			}
			
			#read-host			
			
		}		

# $MyReport | out-file -encoding UTF8 -filepath "\\kaluga.cbr.ru\GU\Inform\Список КО и ОК\Список КО.html"
# $MyReport





if ($icnt1 -gt 0){
     if ($bankArr.count -eq 1){
	$bsort= $bankArr # из-за глюка powershell при сортировке возвращает не массив а строку, если в исходном массиве один элемент.

     }
     else
     {
	$bsort= $bankArr | sort
     }
     $numPP=1
     for ($i=0;$i -lt $bsort.count;$i++){

	$MyReport += Get-HTMLReportBodyShort $numPP 	$bsort[$i]
	$numPP++
     }
}
else
{
       $MyReport +='<tr class="z3"><td colspan="4" >ИНФОРМАЦИЯ ПО ДАННЫМ КРИТЕРИЯМ ВЫБОРКИ ОТСУТСТВУЕТ</td></tr>'
}


$Filename+="($icnt1).html"
$MyReport+=Get-HTMLFooter
$MyReport | out-file -encoding UTF8 -filepath $Filename

Write-output "Program END"
Stop-Transcript