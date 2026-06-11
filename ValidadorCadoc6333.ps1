param(
    [Parameter(Mandatory = $true)]
    [string]$ZipPath
)

$ErrorActionPreference = "Stop"

$ExpectedFiles = @(
    "TARIFAS.TXT",
    "PARTICIP.TXT",
    "CONTATOS.TXT",
    "DESCRICA.TXT",
    "DATABASE.TXT"
)

$Layout = @{
    "TARIFAS.TXT" = @{
        HeaderLength = 32
        DetailLength = 318
        HeaderName   = "TARIFAS "
        HasDetails   = $true
    }
    "PARTICIP.TXT" = @{
        HeaderLength = 32
        DetailLength = 62
        HeaderName   = "PARTICIP"
        HasDetails   = $true
    }
    "CONTATOS.TXT" = @{
        HeaderLength = 32
        DetailLength = 334
        HeaderName   = "CONTATOS"
        HasDetails   = $true
    }
    "DESCRICA.TXT" = @{
        HeaderLength = 32
        DetailLength = 1008
        HeaderName   = "DESCRICA"
        HasDetails   = $true
    }
    "DATABASE.TXT" = @{
        HeaderLength = 30
        DetailLength = 0
        HeaderName   = "DATABASE"
        HasDetails   = $false
    }
}

$Errors = New-Object System.Collections.Generic.List[string]
$Warnings = New-Object System.Collections.Generic.List[string]

function Add-Error {
    param([string]$Message)
    $script:Errors.Add($Message)
}

function Add-Warning {
    param([string]$Message)
    $script:Warnings.Add($Message)
}

function Test-OnlyDigits {
    param([string]$Value)
    return $Value -match '^\d+$'
}

function Get-TextIso88591 {
    param([string]$Path)

    $encoding = [System.Text.Encoding]::GetEncoding("ISO-8859-1")
    return [System.IO.File]::ReadAllText($Path, $encoding)
}

function Split-LinesStrict {
    param([string]$Content)

    $normalized = $Content -replace "`r`n", "`n" -replace "`r", "`n"

    if ($normalized.EndsWith("`n")) {
        Add-Warning "Arquivo termina com quebra de linha. Verifique se isso não gerou linha em branco no final."
        $normalized = $normalized.TrimEnd("`n")
    }

    if ([string]::IsNullOrEmpty($normalized)) {
        return @()
    }

    return $normalized -split "`n"
}

function Validate-HeaderCommon {
    param(
        [string]$FileName,
        [string]$Header,
        [hashtable]$Spec
    )

    if ($Header.Length -ne $Spec.HeaderLength) {
        Add-Error "$FileName: HEADER possui $($Header.Length) caracteres; esperado: $($Spec.HeaderLength)."
        return
    }

    $nomeArquivo = $Header.Substring(0, 8)
    $dataGeracao = $Header.Substring(8, 8)
    $iap = $Header.Substring(16, 8)

    if ($nomeArquivo -ne $Spec.HeaderName) {
        Add-Error "$FileName: nome no HEADER inválido. Encontrado '$nomeArquivo'; esperado '$($Spec.HeaderName)'."
    }

    if (-not (Test-OnlyDigits $dataGeracao)) {
        Add-Error "$FileName: data de geração no HEADER deve conter apenas números AAAAMMDD. Valor: '$dataGeracao'."
    }
    else {
        try {
            [void][datetime]::ParseExact($dataGeracao, "yyyyMMdd", $null)
        }
        catch {
            Add-Error "$FileName: data de geração inválida no HEADER. Valor: '$dataGeracao'."
        }
    }

    if ($iap.Trim().Length -eq 0) {
        Add-Error "$FileName: IAP no HEADER está vazio."
    }

    if ($iap -match '[\.;/\-\s]') {
        Add-Warning "$FileName: IAP contém pontuação ou espaço. Normalmente deve ter 8 caracteres alfanuméricos sem máscara. Valor: '$iap'."
    }

    if ($Spec.HasDetails) {
        $quantidadeRegistros = $Header.Substring(24, 8)

        if (-not (Test-OnlyDigits $quantidadeRegistros)) {
            Add-Error "$FileName: quantidade de registros no HEADER deve conter apenas números. Valor: '$quantidadeRegistros'."
        }
    }
}

function Validate-RecordCount {
    param(
        [string]$FileName,
        [string]$Header,
        [int]$ActualDetailCount
    )

    $declared = $Header.Substring(24, 8)

    if (Test-OnlyDigits $declared) {
        $declaredInt = [int]$declared

        if ($declaredInt -ne $ActualDetailCount) {
            Add-Error "$FileName: quantidade de registros no HEADER é $declaredInt, mas existem $ActualDetailCount linhas de dados."
        }
    }
}

function Validate-NoSeparators {
    param(
        [string]$FileName,
        [string[]]$Lines
    )

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $lineNumber = $i + 1
        $line = $Lines[$i]

        if ($line.Contains(";")) {
            Add-Error "$FileName linha ${lineNumber}: contém ';'. O leiaute é posicional e não deve usar separador."
        }

        if ($line.Contains("`t")) {
            Add-Error "$FileName linha ${lineNumber}: contém TAB. O leiaute é posicional e não deve usar tabulação."
        }

        if ($line.Length -eq 0) {
            Add-Error "$FileName linha ${lineNumber}: linha vazia encontrada."
        }
    }
}

function Validate-TarifasDetail {
    param(
        [string]$Line,
        [int]$LineNumber
    )

    $ano = $Line.Substring(0, 4)
    $trimestre = $Line.Substring(4, 1)
    $proposito = $Line.Substring(5, 1)
    $modalidade = $Line.Substring(6, 1)
    $abrangencia = $Line.Substring(7, 1)

    $tarifaPercentual = $Line.Substring(272, 4)
    $tarifaMonetaria = $Line.Substring(276, 4)
    $tetoMonetario = $Line.Substring(280, 4)
    $tarifaEfetiva = $Line.Substring(284, 4)
    $quantidadeTransacoes = $Line.Substring(288, 12)
    $valorTransacoes = $Line.Substring(300, 15)
    $moeda = $Line.Substring(315, 3)

    if (-not (Test-OnlyDigits $ano)) { Add-Error "TARIFAS.TXT linha ${LineNumber}: Ano inválido: '$ano'." }
    if ($trimestre -notin @("1", "2", "3", "4")) { Add-Error "TARIFAS.TXT linha ${LineNumber}: Trimestre inválido: '$trimestre'." }
    if ($proposito -notin @("1", "2")) { Add-Error "TARIFAS.TXT linha ${LineNumber}: Propósito inválido: '$proposito'." }
    if ($modalidade -notin @("1", "2", "3", "4")) { Add-Error "TARIFAS.TXT linha ${LineNumber}: Modalidade inválida: '$modalidade'." }
    if ($abrangencia -notin @("1", "2")) { Add-Error "TARIFAS.TXT linha ${LineNumber}: Abrangência inválida: '$abrangencia'." }

    foreach ($field in @(
        @{ Name = "Tarifa percentual"; Value = $tarifaPercentual },
        @{ Name = "Tarifa monetária"; Value = $tarifaMonetaria },
        @{ Name = "Teto monetário"; Value = $tetoMonetario },
        @{ Name = "Tarifa efetiva"; Value = $tarifaEfetiva },
        @{ Name = "Quantidade de transações"; Value = $quantidadeTransacoes },
        @{ Name = "Valor das transações"; Value = $valorTransacoes }
    )) {
        if (-not (Test-OnlyDigits $field.Value)) {
            Add-Error "TARIFAS.TXT linha ${LineNumber}: $($field.Name) deve conter apenas números. Valor: '$($field.Value)'."
        }
    }

    if ($moeda.Trim().Length -ne 3) {
        Add-Error "TARIFAS.TXT linha ${LineNumber}: Moeda inválida. Valor: '$moeda'."
    }
}

function Validate-ParticipDetail {
    param(
        [string]$Line,
        [int]$LineNumber
    )

    $ano = $Line.Substring(0, 4)
    $trimestre = $Line.Substring(4, 1)
    $proposito = $Line.Substring(5, 1)
    $modalidade = $Line.Substring(6, 1)
    $abrangencia = $Line.Substring(7, 1)
    $tipoRelacionamento = $Line.Substring(8, 1)
    $participante = $Line.Substring(9, 8)
    $valorTarifas = $Line.Substring(17, 15)
    $quantidadeTransacoes = $Line.Substring(32, 12)
    $valorTransacoes = $Line.Substring(44, 15)
    $moeda = $Line.Substring(59, 3)

    if (-not (Test-OnlyDigits $ano)) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Ano inválido: '$ano'." }
    if ($trimestre -notin @("1", "2", "3", "4")) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Trimestre inválido: '$trimestre'." }
    if ($proposito -notin @("1", "2")) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Propósito inválido: '$proposito'." }
    if ($modalidade -notin @("1", "2", "3", "4")) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Modalidade inválida: '$modalidade'." }
    if ($abrangencia -notin @("1", "2")) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Abrangência inválida: '$abrangencia'." }
    if ($tipoRelacionamento -notin @("1", "2", "3", "4", "5", "6")) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Tipo de relacionamento inválido: '$tipoRelacionamento'." }

    if ($participante.Trim().Length -eq 0) {
        Add-Error "PARTICIP.TXT linha ${LineNumber}: Participante vazio."
    }

    if (-not (Test-OnlyDigits $valorTarifas)) {
        Add-Error "PARTICIP.TXT linha ${LineNumber}: Valor total das tarifas deve conter apenas números. Valor: '$valorTarifas'."
    }

    if (($quantidadeTransacoes.Trim() -ne "ND") -and (-not (Test-OnlyDigits $quantidadeTransacoes))) {
        Add-Error "PARTICIP.TXT linha ${LineNumber}: Quantidade de transações deve ser numérica ou ND. Valor: '$quantidadeTransacoes'."
    }

    if (($valorTransacoes.Trim() -ne "ND") -and (-not (Test-OnlyDigits $valorTransacoes))) {
        Add-Error "PARTICIP.TXT linha ${LineNumber}: Valor das transações deve ser numérico ou ND. Valor: '$valorTransacoes'."
    }

    if ($moeda.Trim().Length -ne 3) {
        Add-Error "PARTICIP.TXT linha ${LineNumber}: Moeda inválida. Valor: '$moeda'."
    }
}

function Validate-ContatosDetail {
    param(
        [string]$Line,
        [int]$LineNumber
    )

    $ano = $Line.Substring(0, 4)
    $trimestre = $Line.Substring(4, 1)
    $tipoContato = $Line.Substring(5, 1)
    $email = $Line.Substring(156, 50).Trim()

    if (-not (Test-OnlyDigits $ano)) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Ano inválido: '$ano'." }
    if ($trimestre -notin @("1", "2", "3", "4")) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Trimestre inválido: '$trimestre'." }
    if ($tipoContato -notin @("D", "T", "I")) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Tipo de contato inválido: '$tipoContato'. Use D, T ou I." }

    if ($email.Length -eq 0) {
        Add-Error "CONTATOS.TXT linha ${LineNumber}: E-mail vazio."
    }
    elseif ($email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        Add-Warning "CONTATOS.TXT linha ${LineNumber}: E-mail aparenta estar fora do padrão: '$email'."
    }
}

function Validate-DescricaDetail {
    param(
        [string]$Line,
        [int]$LineNumber
    )

    $ano = $Line.Substring(0, 4)
    $trimestre = $Line.Substring(4, 1)
    $proposito = $Line.Substring(5, 1)
    $modalidade = $Line.Substring(6, 1)
    $abrangencia = $Line.Substring(7, 1)
    $descricao = $Line.Substring(8, 1000).Trim()

    if (-not (Test-OnlyDigits $ano)) { Add-Error "DESCRICA.TXT linha ${LineNumber}: Ano inválido: '$ano'." }
    if ($trimestre -notin @("1", "2", "3", "4")) { Add-Error "DESCRICA.TXT linha ${LineNumber}: Trimestre inválido: '$trimestre'." }
    if ($proposito -notin @("1", "2")) { Add-Error "DESCRICA.TXT linha ${LineNumber}: Propósito inválido: '$proposito'." }
    if ($modalidade -notin @("1", "2", "3", "4")) { Add-Error "DESCRICA.TXT linha ${LineNumber}: Modalidade inválida: '$modalidade'." }
    if ($abrangencia -notin @("1", "2")) { Add-Error "DESCRICA.TXT linha ${LineNumber}: Abrangência inválida: '$abrangencia'." }

    if ($descricao.Length -eq 0) {
        Add-Error "DESCRICA.TXT linha ${LineNumber}: descrição está vazia."
    }
}

function Validate-Database {
    param(
        [string]$Line
    )

    if ($Line.Length -ne 30) {
        Add-Error "DATABASE.TXT: linha única possui $($Line.Length) caracteres; esperado: 30."
        return
    }

    $nomeArquivo = $Line.Substring(0, 8)
    $dataGeracao = $Line.Substring(8, 8)
    $iap = $Line.Substring(16, 8)
    $dataBase = $Line.Substring(24, 6)

    if ($nomeArquivo -ne "DATABASE") {
        Add-Error "DATABASE.TXT: nome no HEADER inválido. Encontrado '$nomeArquivo'; esperado 'DATABASE'."
    }

    if (-not (Test-OnlyDigits $dataGeracao)) {
        Add-Error "DATABASE.TXT: data de geração deve estar em AAAAMMDD. Valor: '$dataGeracao'."
    }

    if ($iap.Trim().Length -eq 0) {
        Add-Error "DATABASE.TXT: IAP vazio."
    }

    if (-not (Test-OnlyDigits $dataBase)) {
        Add-Error "DATABASE.TXT: data-base deve ser numérica AAAAMM. Valor: '$dataBase'."
        return
    }

    $mes = $dataBase.Substring(4, 2)

    if ($mes -notin @("03", "06", "09", "12")) {
        Add-Error "DATABASE.TXT: data-base inválida '$dataBase'. O mês deve ser 03, 06, 09 ou 12."
    }
}

if (-not (Test-Path $ZipPath)) {
    throw "Arquivo não encontrado: $ZipPath"
}

if ([System.IO.Path]::GetFileName($ZipPath).ToUpperInvariant() -ne "BACEN.ZIP") {
    Add-Warning "O arquivo ZIP deveria se chamar BACEN.ZIP. Nome atual: '$([System.IO.Path]::GetFileName($ZipPath))'."
}

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("CADOC6333_" + [Guid]::NewGuid().ToString("N"))

try {
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)

    try {
        $entries = $zip.Entries

        foreach ($entry in $entries) {
            if ($entry.FullName.EndsWith("/")) {
                Add-Error "O ZIP contém pasta/diretório: '$($entry.FullName)'. Os arquivos devem estar na raiz."
                continue
            }

            if ($entry.FullName -ne $entry.Name) {
                Add-Error "O arquivo '$($entry.FullName)' está dentro de uma pasta. Os arquivos devem estar na raiz do ZIP."
            }
        }

        $actualFiles = $entries | Where-Object { -not $_.FullName.EndsWith("/") } | ForEach-Object { $_.Name.ToUpperInvariant() }

        foreach ($expected in $ExpectedFiles) {
            if ($actualFiles -notcontains $expected) {
                Add-Error "Arquivo obrigatório ausente no ZIP: $expected"
            }
        }

        foreach ($actual in $actualFiles) {
            if ($ExpectedFiles -notcontains $actual) {
                Add-Error "Arquivo extra ou nome incorreto encontrado no ZIP: $actual"
            }
        }

        foreach ($entry in $entries) {
            if ($entry.FullName.EndsWith("/")) {
                continue
            }

            $target = Join-Path $tempDir $entry.Name
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, $true)
        }
    }
    finally {
        $zip.Dispose()
    }

    foreach ($fileName in $ExpectedFiles) {
        $filePath = Join-Path $tempDir $fileName

        if (-not (Test-Path $filePath)) {
            continue
        }

        $content = Get-TextIso88591 $filePath
        $lines = Split-LinesStrict $content

        if ($lines.Count -eq 0) {
            Add-Error "$fileName: arquivo vazio."
            continue
        }

        Validate-NoSeparators $fileName $lines

        $spec = $Layout[$fileName]

        if ($fileName -eq "DATABASE.TXT") {
            if ($lines.Count -ne 1) {
                Add-Error "DATABASE.TXT deve possuir somente uma linha. Encontradas: $($lines.Count)."
            }

            Validate-Database $lines[0]
            continue
        }

        $header = $lines[0]
        $details = @()

        if ($lines.Count -gt 1) {
            $details = $lines[1..($lines.Count - 1)]
        }

        Validate-HeaderCommon $fileName $header $spec
        Validate-RecordCount $fileName $header $details.Count

        for ($i = 0; $i -lt $details.Count; $i++) {
            $line = $details[$i]
            $lineNumber = $i + 2

            if ($line.Length -ne $spec.DetailLength) {
                Add-Error "$fileName linha ${lineNumber}: possui $($line.Length) caracteres; esperado: $($spec.DetailLength)."
                continue
            }

            switch ($fileName) {
                "TARIFAS.TXT"  { Validate-TarifasDetail $line $lineNumber }
                "PARTICIP.TXT" { Validate-ParticipDetail $line $lineNumber }
                "CONTATOS.TXT" { Validate-ContatosDetail $line $lineNumber }
                "DESCRICA.TXT" { Validate-DescricaDetail $line $lineNumber }
            }
        }
    }

    Write-Host ""
    Write-Host "=============================="
    Write-Host "VALIDAÇÃO CADOC 6333"
    Write-Host "=============================="
    Write-Host ""

    if ($Warnings.Count -gt 0) {
        Write-Host "AVISOS:" -ForegroundColor Yellow
        foreach ($warning in $Warnings) {
            Write-Host " - $warning" -ForegroundColor Yellow
        }
        Write-Host ""
    }

    if ($Errors.Count -gt 0) {
        Write-Host "ERROS:" -ForegroundColor Red
        foreach ($errorItem in $Errors) {
            Write-Host " - $errorItem" -ForegroundColor Red
        }

        Write-Host ""
        Write-Host "Resultado: REPROVADO" -ForegroundColor Red
        exit 1
    }
    else {
        Write-Host "Nenhum erro estrutural encontrado." -ForegroundColor Green
        Write-Host "Resultado: APROVADO NA VALIDAÇÃO LOCAL" -ForegroundColor Green
        Write-Host ""
        Write-Host "Observação: esta validação local não substitui a validação oficial do BCB/STA/CRD."
        exit 0
    }
}
finally {
    if (Test-Path $tempDir) {
        Remove-Item $tempDir -Recurse -Force
    }
}