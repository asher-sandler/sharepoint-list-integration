##############################################################
# 
# XML список формируется на основе 
# списка MS Sharepoint 
# http://www.kaluga.cbr.ru/deprts/nadz/Lists/List/view1.aspx
# c фильтром по полям
# ([Тип кредитной организации]=="Банки" и 
# [Регион]=="Филиалы") или ([Регион] == "Калужские банки")
# т.е. выбираем Калужские банки и Филиалы торонних Банков
# 
# 
# Заказчик: Отделы ГУ
# Исполнитель: Астахов А.Б.
# Начато: 26.07.2010
# 
##############################################################

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
Function Get-STName ($id, $bArr){

    $cRet = ""
    for ($jj=0;$jj -lt $bArr.length;$jj++){
    
        if ($bArr[$jj].Split(";")[0] -eq $id){
        
             $cRet = $bArr[$jj].Split(";")[1]
             break
        
        }
    
    }

    Return $cRet

}
#-------------------------------------------------------------------------------------
Function GET-Dolgn ($RukField){
$outStr=""


# write-host $RukField
# берем первую строку заключенную в <b></b>
$RukField = $RukField.Replace("strong","b")
$RF = $RukField.ToLower()

# Write-Host $RukField
$fPos = 0
$lPos = 0

$fPos = $RF.IndexOf("<b>")
$lPos = $RF.IndexOf("</b>")

$OutStr += Del-HTMLMarkup $RukField.Substring($fPos+3,$lPos-$fPos-3)
$OutStr += ";"
$fam1 =  $RF.Substring($lPos,$RukField.Length-$lPos-1)
$fam2 =  $RukField.Substring($lPos,$RukField.Length-$lPos-1)

$fPos = $fam1.IndexOf("<b>")
# write-host $fpos
if ($fPos -le 0){
     $fPos = $fam1.length-1

}
$fam  = Del-HTMLMarkup $fam2.Substring(0,$fPos)

#write-host $fam

$outStr += $fam

#write-host $OutStr
#<p><b>Председатель правления</b>
Return $OutStr
}
#-------------------------------------------------------------------------------------
Function GET-AUTOMATION( $NaimBankST, $NaimGolov, $NaimPiK, $RukDolgn, $RukFam, $UAddress, $TelTbl, $bnkRekv){
$OutStr = ""
#$verb = '<verb>' + $NaimBankST +'</verb>'
if ($NaimGolov.Trim().Length -eq 0){
	$NaimGolov = "Нет данных"
}
$NaimGolov = $NaimGolov.Replace(";","")
if ($NaimPiK.Trim().Length -eq 0){
	$NaimPiK = "Нет данных"
}
$NaimPiK = $NaimPiK.Replace(";","")
if ($RukDolgn.Trim().Length -eq 0){
	$RukDolgn = "Нет данных"
}

$RukDolgn = $RukDolgn.Replace(";","")
if ($RukFam.Trim().Length -eq 0){
	$RukFam = "Нет данных"
}



$RukFam = $RukFam.Replace(";","")
if ($UAddress.Trim().Length -eq 0){
	$UAddress = "Нет данных"
}

$UAddress = $UAddress.Replace(";","")
if ($TelTbl.Trim().Length -eq 0){
	$TelTbl = "Нет данных"
}

$TelTbl = $TelTbl.Replace(";","")

if ($bnkRekv.Trim().Length -eq 0){
	$bnkRekv = "Нет данных"
}

$bnkRekv = $bnkRekv.Replace(";","")
$ToStr  =  $NaimGolov + ";" + $NaimPiK + ";" + $RukDolgn + ";" + $RukFam + ";" + $UAddress + ";" + $TelTbl + ";" + $bnkRekv
$crlf=[char]13+[char]10
$OutStr +=	'		<Convert id="'+$NaimBankST+'">'

$OutStr +=			 $ToStr  
$OutStr +=	'</Convert>'+$crlf




Return $OutStr
}
#-------------------------------------------------------------------------------------

Start-Transcript ‘f:\AdminDir\SpisokKO\STXML.log’ -force
Write-output "Script ST_KO_makeSTXML.ps1"
$FileName = "f:\AdminDir\SpisokKO\ST.XML"


# $FileName
$AlertEmailStr	= "dl.nadz.alert@kaluga.cbr.ru"
$ListLinkStr	= "http://www.kaluga.cbr.ru/deprts/nadz/Lists/List/DispForm.aspx?ID="
$ListLinkStr1	= "&Source=%2Fdeprts%2Fnadz%2Fdefault%2Easpx"
$bankArr = @()
$filArr = @()

$filArr += "1;кбкалуга"
$filArr += "2;кбэлита"
$filArr += "5;кбгэб"

$filArr += "11;сб8608"
$filArr += "12;сб2670"
$filArr += "13;сб5607"
$filArr += "14;сб5568"
$filArr += "15;сб5600"
$filArr += "16;сб7786"


$filArr += "36;ковнешпромбанк"
$filArr += "25;ковтб" # втб
$filArr += "24;комдм" # мдм
$filArr += "37;комиб" # московский индустриальный банк
$filArr += "21;комсэб" # мосстройэкономбанк
$filArr += "33;кообраз" # образование
$filArr += "34;копсб" # промсбербанк
$filArr += "28;копушкино" # пушкино
$filArr += "18;корайфф" # райффайзенг
$filArr += "26;коросбанк" # росбанк
$filArr += "30;короскапитал" # российский капитал
$filArr += "22;косельхоз" # российский сельскохозяйственный банк
$filArr += "29;корусслав" # русславбанк
$filArr += "31;корусстрой" # русстройбанк
$filArr += "10;косвязь" # связь-банк
$filArr += "27;космп" # Северный морской путь Калуга  
$filArr += "19;костратег" # Стратегия
$filArr += "35;коткб" # ТРАНСКАПИТАЛБАНК
$filArr += "32;котранскред" # ТрансКредитБанк Калуга
$filArr += "23;кофора" # ФОРА-БАНК


$filArr = $filArr | sort




$crlf=[char]13+[char]10

$MyReport = '<?xml version="1.0" encoding="utf-8"?>'+$crlf

$MyReport += '<Document>'+$crlf
$MyReport +=            '	<Name>Реквизиты КО</Name>'+$crlf

$MyReport +=			'	<SmartTagCaption>Интел.Замена</SmartTagCaption>'+$crlf

$MyReport +=			'	<SmartTagCount>1</SmartTagCount>'+$crlf

$MyReport +=			'	<VerbCount>7</VerbCount>'+$crlf

$MyReport +=			'	<Recognizer>'+$crlf
$MyReport += 						'		<Desc>Автозамена</Desc>'+$crlf
$MyReport +=						'		<ProgId>SPSAddOnSTag.Recognizer</ProgId>'+$crlf
$MyReport +=						'		<SmartTagDownloadURL></SmartTagDownloadURL>'+$crlf
$MyReport +=						'		<SmartTagName>st#automate</SmartTagName>'+$crlf
$MyReport +=			'	</Recognizer>'+$crlf
$MyReport +=			'	<Action>'+$crlf
$MyReport +=					'		<ProgId>SPSAddOnSTag.Action</ProgId>'+$crlf
$MyReport +=			'	</Action>'+$crlf
$MyReport +=			'	<Methods>'+$crlf
$MyReport +=					'		<Method id="1">'+$crlf
$MyReport +=							'			<Caption>Наим. головной КО</Caption>'+$crlf
$MyReport +=							'			<Function>ConvertGOL</Function>'+$crlf
$MyReport +=					'		</Method>'+$crlf
$MyReport +=					'		<Method id="2">'+$crlf
$MyReport +=							'			<Caption>Наим. филиала</Caption>'+$crlf
$MyReport +=							'			<Function>ConvertFilName</Function>'+$crlf
$MyReport +=					'		</Method>'+$crlf
$MyReport +=					'		<Method id="3">'+$crlf
$MyReport +=							'			<Caption>Должн. лицо</Caption>'+$crlf
$MyReport +=							'			<Function>ConvertBoss</Function>'+$crlf
$MyReport +=					'		</Method>'+$crlf
$MyReport +=					'		<Method id="4">'+$crlf
$MyReport +=							'			<Caption>Фам. рук.</Caption>'+$crlf
$MyReport +=							'			<Function>ConvertBossFam</Function>'+$crlf
$MyReport +=					'		</Method>'+$crlf
$MyReport +=					'		<Method id="5">'+$crlf
$MyReport +=							'			<Caption>Адрес</Caption>'+$crlf
$MyReport +=							'			<Function>ConvertAdress</Function>'+$crlf
$MyReport +=					'		</Method>'+$crlf
$MyReport +=					'		<Method id="6">'+$crlf
$MyReport +=							'			<Caption>Телефон</Caption>'+$crlf
$MyReport +=							'			<Function>ConvertPhone</Function>'+$crlf
$MyReport +=					'		</Method>'+$crlf
$MyReport +=					'		<Method id="7">'+$crlf
$MyReport +=							'			<Caption>Плат. реквизиты</Caption>'+$crlf
$MyReport +=							'			<Function>ConvertPlat</Function>'+$crlf
$MyReport +=					'		</Method>'+$crlf
$MyReport +=			'	</Methods>'+$crlf


$MyReport += '	<Automation>'+$crlf

#$MyReport

		$env:SPpath = "${env:CommonProgramFiles}\Microsoft Shared\web server extensions\12\"
		[System.Reflection.Assembly]::LoadFrom("$env:SPPath\ISAPI\Microsoft.SharePoint.dll") 
        # write-host open web
		# открываем web
		$nsite="http://www.kaluga.cbr.ru/deprts/nadz/"
		$SpSite = New-Object -TypeName "Microsoft.SharePoint.SPSite" -ArgumentList $nsite;
		$spweb=$spsite.OpenWeb();
        # write-host open Sharepoint list
		# открываем  список
		$nlist="http://www.kaluga.cbr.ru/deprts/nadz/Lists/List/view1.aspx"
		$splist=$spweb.getlist($nlist);
		$iCnt = $splist.Items.Count;
		
		# $icnt;
		$icnt1=$icnt2=0;
		
		$bnkCount = 0
		$VerbsStr = ""
		
		
		
		
		
		
		for ($jj=0; $jj -lt $iCnt; $jj++){
			
			$spcurItem = $spList.Items.item($jj);
			# Калужские банки
			
			
			
			
			#$spcurItem["Банк"]
			#$spcurItem["Регион"]
			#$spcurItem["Тип кредитной организации"]
			
			
			
			$HTMLrowString = ""
			
			$RukTbl = $NiD_Dov = $Srok_Dov = $Srok_Dov2=  $PAddress = $Ogrn = $OgrnDate = $KS4et = $Bik = $Okved = ""
			$ObslRKC = $Okpo = $Inn = $Kpp= $TelTbl = $UAddress = $NaimBankST = $NaimPiK = $NaimGolov = $RukDolgn = $RukFam = $bnkRekv = ""
			
			$bnkID		= 0
			$bnkID		= $spcurItem["ИД"]
			
			
			$NaimGolov	= Del-HTMLMarkup $spcurItem["Наименование головной кредитной организации, адрес"]
			$NaimPiK	= Del-HTMLMarkup $spcurItem["Наименование (полное и краткое)"]
			$NaimGolov  = $NaimGolov.replace($crlf,'')
			$NaimPiK    = $NaimPiK.replace($crlf,'')
			$NaimGolov  = $NaimGolov.replace('&quot;','"')
			$NaimPiK    = $NaimPiK.replace('&quot;','"')
			$NaimGolov  = $NaimGolov.replace('&nbsp;',' ')
			$NaimPiK    = $NaimPiK.replace('&nbsp;',' ')
			
			
			
			if ($spcurItem["Телефон"].length -gt 0){
				$TelTbl		= Del-HTMLMarkup $spcurItem["Телефон"]
				}
			else{
				$TelTbl = ""
				}	
			$UAddress	= Del-HTMLMarkup $spcurItem["Юридический адрес"]
			$UAddress	= $UAddress.Trim()
			$PAddress	= Del-HTMLMarkup $spcurItem["Фактический адрес"]
			$Okpo		= $spcurItem["ОКПО"]
			$Inn		= $spcurItem["ИНН"] 
			$Kpp		= $spcurItem["КПП"]
			$Ogrn		= $spcurItem["ОГРН"]
			$OgrnDate	= Del-HTMLMarkup $spcurItem["серия, №, дата свидетельства"]
			$KS4et		= Del-HTMLMarkup $spcurItem["к/счет (субсчет), наименование РКЦ"]
			$Bik		= $spcurItem["БИК"]
			$Okved		= $spcurItem["ОКВЭД"]
			$ObslRKC    = $spcurItem["Обслуживается в РКЦ"]
			
			
			$bnkRekv = ""
			if ($Bik.length -gt 0){
			    $bnkRekv += "БИК "  + $Bik.Trim() 
			}
			if ($Inn -gt 0){
				if ($bnkRekv.length -gt 0){
					$bnkRekv += ", "
			    }
				$bnkRekv += "ИНН "  + $Inn 
			}
			if ($Kpp -gt 0){	
				if ($bnkRekv.length -gt 0){
					$bnkRekv += ", "
			    }
				$bnkRekv += "КПП "  + $Kpp 
			}
			if ($Okpo.length -gt 0){	
				if ($bnkRekv.length -gt 0){
					$bnkRekv += ", "
			    }
				$bnkRekv += "ОКПО " + $($Okpo.Replace(" ","")).trim() 
			}
			if ($Okved.length -gt 0){
			    if ($bnkRekv.length -gt 0){
					$bnkRekv += ", "
			    }	
				$bnkRekv += "ОКВЭД "+ $($Okved.Replace(" ","")).trim()  
			}
			if 	(!($ObslRKC -eq "Не обслуживается")){
			    $RkcNaim = ""
			    If  ($ObslRKC -eq "ГРКЦ"){
					$RkcNaim = " в ГРКЦ  Банка России по Калужской области"
			    
			    }
			    If  ($ObslRKC -eq "РКЦ г. Обнинск"){
					$RkcNaim = " в РКЦ Обнинск Банка России по Калужской области"
			    
			    }
			    If  ($ObslRKC -eq "РКЦ г. Киров"){
					$RkcNaim = " в РКЦ Киров Банка России по Калужской области"
			    
			    }
				if ($KS4et.length -gt 20){
					if ($bnkRekv.length -gt 0){
						$bnkRekv += ", "
					}
			    
					$bnkRekv += "К/с № "+ $KS4et.substring(0,20) + $RkcNaim
					$bnkRekv = $bnkRekv.replace("&nbsp;"," ")
				}
			}
			
				
			 
			if (($spcurItem["Регион"] -eq "Калужские банки") -and ($spcurItem["Головной банк"] -eq "Да") ){
				$bnkCount++
				$NaimBankST = Get-STName $bnkID $filArr
			
				if ($VerbsStr.length -gt 0){
				    $VerbsStr+=","
				
				}
				$VerbsStr+=$NaimBankST
				
				$RukDolgn = GET-Dolgn $spcurItem["Руководитель"]
				$RukFam = $RukDolgn.Split(";")[1]
				
				$RukDolgn = $RukDolgn.Split(";")[0]
				
				
				
				$MyReport += GET-AUTOMATION $NaimBankST $NaimGolov $NaimPiK $RukDolgn $RukFam $UAddress $TelTbl $bnkRekv
				
				
			}
			
			if ( ($spcurItem["Тип кредитной организации"] -eq "Банки") -and ($spcurItem["Регион"] -eq "Филиалы") ){
				$bnkCount++
				$NaimBankST = Get-STName $bnkID $filArr
				
				if ($VerbsStr.length -gt 0){
				    $VerbsStr+=","
				
				}
				$VerbsStr+=$NaimBankST
				
				
				$RukDolgn = GET-Dolgn $spcurItem["Руководитель"]
				$RukFam = $RukDolgn.Split(";")[1]
				
				$RukDolgn = $RukDolgn.Split(";")[0]
				
				
				$MyReport += GET-AUTOMATION $NaimBankST $NaimGolov $NaimPiK $RukDolgn $RukFam $UAddress $TelTbl $bnkRekv
				
			}
			
			
		}		

$MyReport += '	</Automation>'+$crlf
$MyReport += '	<Verbs>'+$VerbsStr

$MyReport += '</Verbs>'+$crlf
$MyReport += '	<WhatArrCount>'
$MyReport += $bnkCount
$MyReport += '</WhatArrCount>'+$crlf
$MyReport += '</Document>'+$crlf

$MyReport | out-file -encoding UTF8 -filepath $Filename

Write-output "Script END"
Stop-Transcript