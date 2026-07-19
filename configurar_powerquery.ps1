# Configura o Power Query na Calculadora_OMS.xlsx para puxar o banco automaticamente.
# Rode SEMPRE depois de (re)gerar a planilha com: python build_planilha.py
# Uso:  powershell -ExecutionPolicy Bypass -File configurar_powerquery.ps1

$Calc = Join-Path $PSScriptRoot "Calculadora_OMS.xlsx"
if (-not (Test-Path $Calc)) { Write-Output "Nao achei Calculadora_OMS.xlsx. Rode antes: python build_planilha.py"; return }

# Consulta M: le o caminho do banco da celula 'caminho_bd' (mesma pasta) e importa a aba Dados
$M = @'
let
    caminho = Excel.CurrentWorkbook(){[Name="caminho_bd"]}[Content]{0}[Column1],
    Fonte = Excel.Workbook(File.Contents(caminho), true),
    Dados = Fonte{[Item="Dados",Kind="Sheet"]}[Data],
    Selecionado = Table.SelectColumns(Dados, {"ID / Prontuário","Nome","Sexo","Data nasc.","Data da medida","Peso (kg)","Altura (cm)","Medida"}),
    Renomeado = Table.RenameColumns(Selecionado, {{"ID / Prontuário","ID"},{"Data nasc.","DN"},{"Data da medida","Data"},{"Peso (kg)","Peso"},{"Altura (cm)","Altura"}}),
    Filtrado = Table.SelectRows(Renomeado, each [ID] <> null and [ID] <> "")
in
    Filtrado
'@

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Calc)
    # limpeza idempotente na ordem correta: tabela -> conexao -> consulta
    $ws = $wb.Sheets.Item("Banco"); $ws.Visible = -1
    try { foreach ($lo0 in @($ws.ListObjects)) { $lo0.Delete() } } catch {}
    $ws.Cells.Clear()
    try { foreach ($cn in @($wb.Connections)){ if ($cn.Name -like "*Banco*"){ $cn.Delete() } } } catch {}
    try { foreach ($q in @($wb.Queries))     { if ($q.Name -eq "Banco")    { $q.Delete() } } } catch {}
    $xl.CalculateFull()
    $wb.Queries.Add("Banco", $M) | Out-Null
    $conn = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Banco;Extended Properties=""'
    $lo = $ws.ListObjects.Add(0, $conn, $true, 1, $ws.Range("A1"))
    $qt = $lo.QueryTable
    $qt.CommandType = 2; $qt.CommandText = "SELECT * FROM [Banco]"
    $qt.BackgroundQuery = $false; $qt.RefreshStyle = 1
    $qt.AdjustColumnWidth = $false; $qt.PreserveColumnInfo = $true
    $qt.RefreshOnFileOpen = $true; $qt.SaveData = $true
    $lo.DisplayName = "tb_Banco"
    $err = ""
    try { $qt.Refresh($false) } catch { $err = $_.Exception.Message }
    if ($err -ne "") { Write-Output "REFRESH_ERRO: $err"; $wb.Close($false); return }
    $n = $lo.ListRows.Count
    $xl.CalculateFullRebuild()
    $ws.Visible = 0
    $wb.Save(); $wb.Close($true)
    Write-Output "Power Query configurado. Banco carregado com $n linha(s)."
} catch {
    Write-Output "ERRO: $($_.Exception.Message)"
    try { $wb.Close($false) } catch {}
} finally {
    $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
