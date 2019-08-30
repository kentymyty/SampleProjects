## 
## 【Excelから特定の文字列を検索する】
## 
## ■使い方
## Paramを設定する
## ${dir}⇒検索対象のディレクトリを指定「C:\～～\」
## ${before}⇒Excelファイル内で検索する文字列
## ${after}⇒${before}に置換する文字列
## 
## 例）Param(${dir}="C:\Users\user1\Desktop\test\", ${before}="変換前文字列", ${after}="変換後文字列")
## Param(${dir}="", ${before}="hogehoge", ${after}="hogehoge")
## 

#Param(${dir}="", ${before}="hogehoge", ${after}="hogehoge")
Param(${dir}="C:\Users\user1\Desktop\test\", ${before}="変換前文字列", ${after}="変換後文字列")
$LogFileName = "${dir}output.log"

if (${dir} -eq "") {
    ${dir} = Split-Path $myInvocation.MyCommand.Path -Parent
}

if (${before} -eq "") {
    echo "キーワードが指定されていません。"
    return
}

Write-Output "Dir : ${dir}" | Out-File -FilePath $LogFileName -Encoding Default -append
Write-Output "Keyword : ${before}" | Out-File -FilePath $LogFileName -Encoding Default -append

${excel} = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $FALSE

Get-ChildItem "${dir}" -Include "*.xlsx","*.xls","*.xlt" -Recurse -Name | % {
    ${childPath} = $_
    ${wb} = ${excel}.Workbooks.Open("${dir}\${childPath}")
    ${wb}.Worksheets | % {
        ${ws} = $_
        ${wsName} = ${ws}.Name
        ${first} = ${result} = ${ws}.Cells.Find(${before})
        while (${result} -ne $null) {
            # 検索した「ファイル」「シート」「セル」「値」情報を出力
            Write-Output "${childPath}[${wsName}][$(${result}.Row), $(${result}.Column)] : $(${result}.Text)" | Out-File -FilePath $LogFileName -Encoding Default -append
            
            # 文字列置換の実行
            ${replase} = ${ws}.Cells.Item($(${result}.Row), $(${result}.Column)).Text.Replace(${before}, ${after})
            ${ws}.Cells.Item(${result}.Row, ${result}.Column) = ${replase}
            
            # 次の検索
            ${result} = ${ws}.Cells.FindNext(${result})
        }
    }
    # echo "${dir}${childPath}"
    Write-Output "${dir}${childPath}" | Out-File -FilePath $LogFileName -Encoding Default -append
    ${wb}.SaveAs("${dir}${childPath}")
    ${wb}.Close(0)
}

${excel}.Quit()

${ws} = $null
${wb} = $null
${excel} = $null
    [System.GC]::Collect([System.GC]::MaxGeneration)
