# 指定したExcelのシート名を全て取得するPowerShell
# 引数1：ファイル名

# ファイルのパスを取得
$scriptPF = $MyInvocation.MyCommand.Path
$dirPath = Split-Path -Parent $scriptPF

Get-ChildItem $dirPath -Recurse -Include *.xlsx,*.xls | %{ 
    $excelFileName= $_.BaseName
    if (Test-Path $_.FullName) {
        # Excelのオープン
        $objExcelAp = New-Object -ComObject Excel.Application
        $objExcelAp.Visible = $false

        $objExcelBook = $objExcelAp.Workbooks.Open($_.FullName)

        $objExcelBook.Sheets | ForEach-Object {
            ($excelFileName + ',' + $_.Name) | Out-File ("AllSheetsName.txt") -Append -Encoding Default
        }

        # ファイルを閉じてExcelの終了
        $objExcelBook.close($false)
        $objExcelAp.Quit()

        # 各変数を解放（.ブロックでOut-Nullしているのは実行すると「0 0 0」と表示されるため）
        .{
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcelBook)
            $objExcelBook = $null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcelAp)
            $objExcelAp = $null
        } | Out-Null

        # ガベージコレクトの明示実行
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()

    }
}

