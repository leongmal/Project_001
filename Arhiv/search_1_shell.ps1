$Path = "\\Esk-node2\ора\Договора Стачек\Договора арендаторы"
#$Path = "\\Esk-node2\ора\Договора Стачек\Договора дочерних потребителей"
$Phrase = [regex]::Escape("до начала расчетного месяца, в котором планируется изменение отпускной цены")
Get-ChildItem -Path $Path -Recurse -Include *.doc,*.docx | ForEach-Object {
    try {
        # Попытка открыть как текстовый документ (Word использует UTF-8/Unicode)
        $content = Get-Content -LiteralPath $_.FullName -Encoding Unicode -ErrorAction Stop
    } catch {
        # Если Unicode не подошел (старые .doc), пробуем системную кодировку (обычно cp1251)
        $content = Get-Content -LiteralPath $_.FullName -Encoding Default
    }
    $lines = $content -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match $Phrase) {
            Write-Output "$($_.FullName):$($i + 1)"
            break
        }
    }
}