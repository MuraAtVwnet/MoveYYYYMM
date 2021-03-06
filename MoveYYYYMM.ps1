############################################################
# ファイルを年月フォルダーに移動する
############################################################
Param([switch] $WhatIf)

$GC_LogDir = "C:\Logs"
$GC_LogNmame = "MoveYYYYMM"

# これより古いファイルは処理しない
[datetime]$GC_MinDate = "1990/1/1"

$GC_CurrentDir = (Get-Location).Path


##########################################################################
# ログ出力
##########################################################################
function Log(
			$LogString
			){

	# ログの出力先
	$LogPath = $GC_LogDir

	# ログファイル名
	$LogName = $GC_LogNmame

	$Now = Get-Date

	# Log 出力文字列に時刻を付加(YYYY/MM/DD HH:MM:SS.MMM $LogString)
	$Log = $Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " "
	$Log += $LogString

	# ログファイル名が設定されていなかったらデフォルトのログファイル名をつける
	if( $LogName -eq $null ){
		$LogName = "LOG"
	}

	# ログファイル名(XXXX_YYYY-MM-DD.log)
	$LogFile = $LogName + "_" +$Now.ToString("yyyy-MM-dd") + ".log"

	# ログフォルダーがなかったら作成
	if( -not (Test-Path $LogPath) ) {
		New-Item $LogPath -Type Directory
	}

	# ログファイル名
	$LogFileName = Join-Path $LogPath $LogFile

	# ログ出力
	Write-Output $Log | Out-File -FilePath $LogFileName -Encoding utf8 -append

	# echo させるために出力したログを戻す
	Return $Log
}


##########################################################################
# main
##########################################################################

$LogString = "========= Start ========="
Log $LogString

$LogString = "Current Directory : $GC_CurrentDir"
Log $LogString


[array]$TergetFiles = Get-ChildItem -Path $GC_CurrentDir

$LogString = "Terget Files : " + $TergetFiles.Count
Log $LogString

foreach($TergetFile in $TergetFiles){
	# ディレクトリ
	if( ($TergetFile.Attributes -band [System.IO.FileAttributes]::Directory) -eq [System.IO.FileAttributes]::Directory){
		$LogString = "[WARNING]This is Directory : " + $TergetFile.FullName
		Log $LogString
		continue
	}

	Try{
		[datetime]$FileDateTime = $TergetFile.LastWriteTime
	}
	Catch{
		$LogString = "[ERROR]File Date is Error : $TergetFiles.FullName"
		Log $LogString
		continue
	}

	if( $FileDateTime -lt $GC_MinDate ){
		$LogString = "[ERROR]File Date so small : $FileDateTime.ToString()"
		Log $LogString
		continue
	}

	$TergetYYYYDir = Join-Path $GC_CurrentDir $FileDateTime.ToString("yyyy")
	if(-not (Test-Path $TergetYYYYDir )){
		$LogString = "Make Directory : $TergetYYYYDir"
		Log $LogString
		if(-not $WhatIf){
			mkdir $TergetYYYYDir
		}
	}

	$TergetYYYYMMDir = Join-Path $TergetYYYYDir $FileDateTime.ToString("MM")
	if(-not (Test-Path $TergetYYYYMMDir )){
		$LogString = "Make Directory : $TergetYYYYMMDir"
		Log $LogString
		if(-not $WhatIf){
			mkdir $TergetYYYYMMDir
		}
	}

	$MoveTergetFullPath = Join-Path $TergetYYYYMMDir $TergetFile.Name
	if( Test-Path $MoveTergetFullPath ){
		$LogString = "[ERROR]Terget file is exist : $MoveTergetFullPath"
		Log $LogString
		continue
	}

	$LogString = "File move : " + $TergetFile.FullName + " -> $TergetYYYYMMDir"
	Log $LogString

	if(-not $WhatIf){
		Move-Item -Path $TergetFile.FullName -Destination $TergetYYYYMMDir
	}
}

$LogString = "========= END ========="
Log $LogString
