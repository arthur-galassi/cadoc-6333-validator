param(
    [Parameter(Mandatory = $true)]
    [string]$ZipPath,

    [Parameter(Mandatory = $false)]
    [switch]$ArranjoFechado
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
    "TARIFAS.TXT" = @{ HeaderLength = 32; DetailLength = 318; HeaderName = "TARIFAS "; HasDetails = $true }
    "PARTICIP.TXT" = @{ HeaderLength = 32; DetailLength = 62; HeaderName = "PARTICIP"; HasDetails = $true }
    "CONTATOS.TXT" = @{ HeaderLength = 32; DetailLength = 334; HeaderName = "CONTATOS"; HasDetails = $true }
    "DESCRICA.TXT" = @{ HeaderLength = 32; DetailLength = 1008; HeaderName = "DESCRICA"; HasDetails = $true }
    "DATABASE.TXT" = @{ HeaderLength = 30; DetailLength = 0; HeaderName = "DATABASE"; HasDetails = $false }
}

$Errors = New-Object System.Collections.Generic.List[string]
$Warnings = New-Object System.Collections.Generic.List[string]

$script:DatabaseDataBase = $null
$script:DatabaseAno = $null
$script:DatabaseTrimestre = $null
$script:HeaderIapByFile = @{}
$script:TarifasKeys = @{}
$script:ParticipKeys = @{}
$script:DescricaKeys = @{}
$script:ContatosCounts = @{ D = 0; T = 0; I = 0 }

function Add-Error { param([string]$Message) $script:Errors.Add($Message) }
function Add-Warning { param([string]$Message) $script:Warnings.Add($Message) }

function Test-OnlyDigits { param([string]$Value) return $Value -match '^\d+$' }
function Test-AlphaNum8 { param([string]$Value) return $Value -cmatch '^[A-Za-z0-9]{8}$' }
function Test-UpperAlpha3 { param([string]$Value) return $Value -cmatch '^[A-Z]{3}$' }
function Test-AllZeros { param([string]$Value) return $Value -match '^0+$' }

function Test-ValidDateYyyyMMdd {
    param([string]$Value)
    if (-not (Test-OnlyDigits $Value)) { return $false }
    try { [void][datetime]::ParseExact($Value, "yyyyMMdd", [Globalization.CultureInfo]::InvariantCulture); return $true }
    catch { return $false }
}

function Get-TextIso88591 {
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        Add-Error "$([System.IO.Path]::GetFileName($Path)): arquivo contém BOM UTF-8. A codificação exigida é ISO-8859-1."
    }

    if ($bytes.Length -ge 2 -and (($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) -or ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF))) {
        Add-Error "$([System.IO.Path]::GetFileName($Path)): arquivo contém BOM UTF-16. A codificação exigida é ISO-8859-1."
    }

    $encoding = [System.Text.Encoding]::GetEncoding("ISO-8859-1")
    return $encoding.GetString($bytes)
}

function Split-LinesStrict {
    param([string]$Content)

    $normalized = $Content -replace "`r`n", "`n" -replace "`r", "`n"

    if ($normalized.EndsWith("`n")) {
        Add-Warning "Arquivo termina com quebra de linha. Verifique se isso não gerou linha em branco no final."
        $normalized = $normalized.TrimEnd("`n")
    }

    if ([string]::IsNullOrEmpty($normalized)) { return @() }
    return $normalized -split "`n"
}

function Get-QuarterFromMonth {
    param([string]$Month)
    switch ($Month) {
        "03" { return "1" }
        "06" { return "2" }
        "09" { return "3" }
        "12" { return "4" }
        default { return $null }
    }
}

function Validate-PeriodMatchesDatabase {
    param([string]$FileName, [int]$LineNumber, [string]$Ano, [string]$Trimestre)

    if ($null -eq $script:DatabaseAno -or $null -eq $script:DatabaseTrimestre) { return }
    if ((Test-OnlyDigits $Ano) -and ($Trimestre -in @("1", "2", "3", "4"))) {
        if ($Ano -ne $script:DatabaseAno -or $Trimestre -ne $script:DatabaseTrimestre) {
            Add-Error "$FileName linha ${LineNumber}: Ano/Trimestre ($Ano/$Trimestre) não corresponde à data-base do DATABASE.TXT ($($script:DatabaseDataBase))."
        }
    }
}

function Validate-NumericFixed {
    param([string]$FileName, [int]$LineNumber, [string]$FieldName, [string]$Value, [int]$Length, [switch]$AllowND)

    if ($AllowND -and $Value -eq ("ND" + (" " * ($Length - 2)))) { return }

    if ($AllowND -and $Value.Trim() -eq "ND") {
        Add-Error "$FileName linha ${LineNumber}: $FieldName com 'ND' deve estar alinhado à esquerda e preenchido com espaços à direita até $Length caracteres. Valor bruto: '$Value'."
        return
    }

    if ($Value.Length -ne $Length -or -not (Test-OnlyDigits $Value)) {
        $extra = if ($AllowND) { " ou 'ND' alinhado à esquerda e preenchido com espaços" } else { "" }
        Add-Error "$FileName linha ${LineNumber}: $FieldName deve conter exatamente $Length dígitos, com zeros à esquerda quando necessário$extra. Valor: '$Value'."
    }
}

function Validate-TextFixed {
    param([string]$FileName, [int]$LineNumber, [string]$FieldName, [string]$Value, [int]$Length, [switch]$AllowNA)

    if ($Value.Length -ne $Length) {
        Add-Error "$FileName linha ${LineNumber}: $FieldName deve possuir $Length caracteres. Valor possui $($Value.Length)."
        return
    }

    if ($Value -match '[\x00-\x08\x0B\x0C\x0E-\x1F]') {
        Add-Error "$FileName linha ${LineNumber}: $FieldName contém caractere de controle inválido."
    }

    $trimRight = $Value.TrimEnd(' ')
    $trimBoth = $Value.Trim()

    if ($trimBoth.Length -eq 0) {
        Add-Error "$FileName linha ${LineNumber}: $FieldName está em branco. Use texto válido ou 'NA' quando a dimensão não for aplicável."
        return
    }

    if ($Value.StartsWith(" ")) {
        Add-Error "$FileName linha ${LineNumber}: $FieldName deve ser preenchido da esquerda para a direita, sem espaços à esquerda. Valor: '$Value'."
    }

    if ($trimBoth -ne $trimRight) {
        Add-Error "$FileName linha ${LineNumber}: $FieldName deve ter apenas espaços de preenchimento à direita, não à esquerda. Valor: '$Value'."
    }

    if ($AllowNA -and $trimBoth.ToUpperInvariant() -eq "NA" -and $trimBoth -cne "NA") {
        Add-Error "$FileName linha ${LineNumber}: $FieldName usa 'NA' em caixa inválida. Informe exatamente 'NA' quando não aplicável. Valor útil: '$trimBoth'."
    }
}

function Validate-UpperTextFixed {
    param([string]$FileName, [int]$LineNumber, [string]$FieldName, [string]$Value, [int]$Length, [switch]$AllowBlank)

    if ($Value.Length -ne $Length) { Add-Error "$FileName linha ${LineNumber}: $FieldName deve possuir $Length caracteres."; return }
    $useful = $Value.TrimEnd(' ')

    if ($AllowBlank -and $useful.Length -eq 0) { return }
    if ($useful.Length -eq 0) { Add-Error "$FileName linha ${LineNumber}: $FieldName está vazio."; return }
    if ($Value.StartsWith(" ")) { Add-Error "$FileName linha ${LineNumber}: $FieldName deve ser preenchido da esquerda para a direita, sem espaços à esquerda." }
    if ($useful -match '[a-zà-öø-ÿ]') { Add-Error "$FileName linha ${LineNumber}: $FieldName deve estar em maiúsculas, com espaços apenas à direita. Valor útil: '$useful'." }
    if ($Value -match '[\x00-\x08\x0B\x0C\x0E-\x1F]') { Add-Error "$FileName linha ${LineNumber}: $FieldName contém caractere de controle inválido." }
}

function Validate-HeaderCommon {
    param([string]$FileName, [string]$Header, [hashtable]$Spec)

    if ($Header.Length -ne $Spec.HeaderLength) {
        Add-Error "$FileName: HEADER possui $($Header.Length) caracteres; esperado: $($Spec.HeaderLength)."
        return
    }

    $nomeArquivo = $Header.Substring(0, 8)
    $dataGeracao = $Header.Substring(8, 8)
    $iap = $Header.Substring(16, 8)
    $script:HeaderIapByFile[$FileName] = $iap

    if ($nomeArquivo -ne $Spec.HeaderName) { Add-Error "$FileName: nome no HEADER inválido. Encontrado '$nomeArquivo'; esperado '$($Spec.HeaderName)'." }

    if (-not (Test-ValidDateYyyyMMdd $dataGeracao)) { Add-Error "$FileName: data de geração no HEADER deve ser uma data válida em AAAAMMDD. Valor: '$dataGeracao'." }

    if (-not (Test-AlphaNum8 $iap)) { Add-Error "$FileName: IAP no HEADER deve conter exatamente 8 caracteres alfanuméricos, sem máscara/pontuação/espaços. Valor: '$iap'." }

    if ($Spec.HasDetails) {
        $quantidadeRegistros = $Header.Substring(24, 8)
        if (-not (Test-OnlyDigits $quantidadeRegistros)) { Add-Error "$FileName: quantidade de registros no HEADER deve conter exatamente 8 dígitos. Valor: '$quantidadeRegistros'." }
    }
}

function Validate-RecordCount { param([string]$FileName, [string]$Header, [int]$ActualDetailCount)
    $declared = $Header.Substring(24, 8)
    if (Test-OnlyDigits $declared) {
        $declaredInt = [int]$declared
        if ($declaredInt -ne $ActualDetailCount) { Add-Error "$FileName: quantidade de registros no HEADER é $declaredInt, mas existem $ActualDetailCount linhas de dados." }
    }
}

function Validate-NoSeparators {
    param([string]$FileName, [string[]]$Lines)
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $lineNumber = $i + 1
        $line = $Lines[$i]
        if ($line.Contains(";")) { Add-Error "$FileName linha ${lineNumber}: contém ';'. O leiaute é posicional e não deve usar separador." }
        if ($line.Contains("`t")) { Add-Error "$FileName linha ${lineNumber}: contém TAB. O leiaute é posicional e não deve usar tabulação." }
        if ($line.Length -eq 0 -or $line.Trim().Length -eq 0) { Add-Error "$FileName linha ${lineNumber}: linha vazia ou composta apenas por espaços encontrada." }
    }
}

function Add-DuplicateCheck {
    param([hashtable]$Store, [string]$Key, [string]$FileName, [int]$LineNumber, [string]$Description)
    if ($Store.ContainsKey($Key)) { Add-Error "$FileName linha ${LineNumber}: registro duplicado para $Description. Primeira ocorrência na linha $($Store[$Key])." }
    else { $Store[$Key] = $LineNumber }
}

function Validate-TarifasDetail {
    param([string]$Line, [int]$LineNumber)

    $ano = $Line.Substring(0, 4); $trimestre = $Line.Substring(4, 1); $proposito = $Line.Substring(5, 1); $modalidade = $Line.Substring(6, 1); $abrangencia = $Line.Substring(7, 1)
    $segmento = $Line.Substring(8, 64); $parcelas = $Line.Substring(72, 8); $produto = $Line.Substring(80, 64); $captura = $Line.Substring(144, 64); $natureza = $Line.Substring(208, 64)
    $tarifaPercentual = $Line.Substring(272, 4); $tarifaMonetaria = $Line.Substring(276, 4); $tetoMonetario = $Line.Substring(280, 4); $tarifaEfetiva = $Line.Substring(284, 4)
    $quantidadeTransacoes = $Line.Substring(288, 12); $valorTransacoes = $Line.Substring(300, 15); $moeda = $Line.Substring(315, 3)

    if (-not (Test-OnlyDigits $ano)) { Add-Error "TARIFAS.TXT linha ${LineNumber}: Ano deve conter exatamente 4 dígitos. Valor: '$ano'." }
    if ($trimestre -notin @("1", "2", "3", "4")) { Add-Error "TARIFAS.TXT linha ${LineNumber}: Trimestre inválido: '$trimestre'." }
    if ($proposito -notin @("1", "2")) { Add-Error "TARIFAS.TXT linha ${LineNumber}: Propósito inválido: '$proposito'." }
    if ($modalidade -notin @("1", "2", "3", "4")) { Add-Error "TARIFAS.TXT linha ${LineNumber}: Modalidade inválida: '$modalidade'." }
    if ($abrangencia -notin @("1", "2")) { Add-Error "TARIFAS.TXT linha ${LineNumber}: Abrangência inválida: '$abrangencia'." }

    Validate-PeriodMatchesDatabase "TARIFAS.TXT" $LineNumber $ano $trimestre

    Validate-TextFixed "TARIFAS.TXT" $LineNumber "Segmento" $segmento 64 -AllowNA
    Validate-TextFixed "TARIFAS.TXT" $LineNumber "Número de parcelas" $parcelas 8 -AllowNA
    Validate-TextFixed "TARIFAS.TXT" $LineNumber "Produto" $produto 64 -AllowNA
    Validate-TextFixed "TARIFAS.TXT" $LineNumber "Forma de captura" $captura 64 -AllowNA
    Validate-TextFixed "TARIFAS.TXT" $LineNumber "Natureza do recebedor" $natureza 64 -AllowNA

    Validate-NumericFixed "TARIFAS.TXT" $LineNumber "Tarifa de intercâmbio definida em termos percentuais" $tarifaPercentual 4
    Validate-NumericFixed "TARIFAS.TXT" $LineNumber "Tarifa de intercâmbio definida em valores monetários" $tarifaMonetaria 4
    Validate-NumericFixed "TARIFAS.TXT" $LineNumber "Teto para a tarifa de intercâmbio definida em valores monetários" $tetoMonetario 4
    Validate-NumericFixed "TARIFAS.TXT" $LineNumber "Tarifa de intercâmbio efetiva" $tarifaEfetiva 4
    Validate-NumericFixed "TARIFAS.TXT" $LineNumber "Quantidade de transações" $quantidadeTransacoes 12
    Validate-NumericFixed "TARIFAS.TXT" $LineNumber "Valor das transações" $valorTransacoes 15

    if (-not (Test-UpperAlpha3 $moeda)) { Add-Error "TARIFAS.TXT linha ${LineNumber}: Moeda deve conter exatamente 3 letras maiúsculas ISO 4217. Valor: '$moeda'." }
    if ($abrangencia -eq "1" -and $moeda -ne "BRL") { Add-Error "TARIFAS.TXT linha ${LineNumber}: Arranjos domésticos (abrangência=1) devem reportar moeda BRL. Valor: '$moeda'." }

    if ($ArranjoFechado -and ((-not (Test-AllZeros $tarifaPercentual)) -or (-not (Test-AllZeros $tarifaMonetaria)) -or (-not (Test-AllZeros $tetoMonetario)) -or (-not (Test-AllZeros $tarifaEfetiva)))) {
        Add-Error "TARIFAS.TXT linha ${LineNumber}: para arranjo fechado (-ArranjoFechado), os campos de tarifa de intercâmbio devem ser zero."
    }

    $key = ($Line.Substring(0, 272))
    Add-DuplicateCheck $script:TarifasKeys $key "TARIFAS.TXT" $LineNumber "a mesma combinação de dimensões (Ano, Trimestre, Propósito, Modalidade, Abrangência, Segmento, Parcelas, Produto, Captura e Natureza)"
}

function Validate-ParticipDetail {
    param([string]$Line, [int]$LineNumber)

    $ano = $Line.Substring(0, 4); $trimestre = $Line.Substring(4, 1); $proposito = $Line.Substring(5, 1); $modalidade = $Line.Substring(6, 1); $abrangencia = $Line.Substring(7, 1)
    $tipoRelacionamento = $Line.Substring(8, 1); $participante = $Line.Substring(9, 8); $valorTarifas = $Line.Substring(17, 15); $quantidadeTransacoes = $Line.Substring(32, 12); $valorTransacoes = $Line.Substring(44, 15); $moeda = $Line.Substring(59, 3)

    if (-not (Test-OnlyDigits $ano)) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Ano deve conter exatamente 4 dígitos. Valor: '$ano'." }
    if ($trimestre -notin @("1", "2", "3", "4")) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Trimestre inválido: '$trimestre'." }
    if ($proposito -notin @("1", "2")) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Propósito inválido: '$proposito'." }
    if ($modalidade -notin @("1", "2", "3", "4")) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Modalidade inválida: '$modalidade'." }
    if ($abrangencia -notin @("1", "2")) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Abrangência inválida: '$abrangencia'." }
    if ($tipoRelacionamento -notin @("1", "2", "3", "4", "5", "6")) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Tipo de relacionamento inválido: '$tipoRelacionamento'." }

    Validate-PeriodMatchesDatabase "PARTICIP.TXT" $LineNumber $ano $trimestre

    if (-not (Test-AlphaNum8 $participante)) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Participante deve conter exatamente 8 caracteres alfanuméricos, sem pontos/máscara. Valor: '$participante'." }

    Validate-NumericFixed "PARTICIP.TXT" $LineNumber "Valor total das tarifas cobradas pelo IAP" $valorTarifas 15
    Validate-NumericFixed "PARTICIP.TXT" $LineNumber "Quantidade de transações" $quantidadeTransacoes 12 -AllowND
    Validate-NumericFixed "PARTICIP.TXT" $LineNumber "Valor das transações" $valorTransacoes 15 -AllowND

    if (-not (Test-UpperAlpha3 $moeda)) { Add-Error "PARTICIP.TXT linha ${LineNumber}: Moeda deve conter exatamente 3 letras maiúsculas ISO 4217. Valor: '$moeda'." }
    if ($abrangencia -eq "1" -and $moeda -ne "BRL") { Add-Error "PARTICIP.TXT linha ${LineNumber}: Arranjos domésticos (abrangência=1) devem reportar moeda BRL. Valor: '$moeda'." }

    if ($ArranjoFechado) {
        if (-not (Test-AllZeros $valorTarifas)) { Add-Error "PARTICIP.TXT linha ${LineNumber}: para arranjo fechado (-ArranjoFechado), Valor total das tarifas cobradas pelo IAP deve ser zero." }
        $iap = $script:HeaderIapByFile["PARTICIP.TXT"]
        if ($iap -and $participante -ne $iap) { Add-Warning "PARTICIP.TXT linha ${LineNumber}: para arranjo fechado, o participante normalmente é o próprio IAP informado no HEADER. Participante: '$participante'; IAP: '$iap'." }
    }

    $key = ($Line.Substring(0, 17))
    Add-DuplicateCheck $script:ParticipKeys $key "PARTICIP.TXT" $LineNumber "a mesma combinação de dimensões, tipo de relacionamento e participante"
}

function Validate-ContatosDetail {
    param([string]$Line, [int]$LineNumber)

    $ano = $Line.Substring(0, 4); $trimestre = $Line.Substring(4, 1); $tipoContato = $Line.Substring(5, 1)
    $nomeRaw = $Line.Substring(6, 50); $cargoRaw = $Line.Substring(56, 50); $telefoneRaw = $Line.Substring(106, 50); $emailRaw = $Line.Substring(156, 50); $enderecoRaw = $Line.Substring(206, 128)
    $email = $emailRaw.TrimEnd(' ')

    if (-not (Test-OnlyDigits $ano)) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Ano deve conter exatamente 4 dígitos. Valor: '$ano'." }
    if ($trimestre -notin @("1", "2", "3", "4")) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Trimestre inválido: '$trimestre'." }
    if ($tipoContato -notin @("D", "T", "I")) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Tipo de contato inválido: '$tipoContato'. Use D, T ou I." }
    else { $script:ContatosCounts[$tipoContato] = [int]$script:ContatosCounts[$tipoContato] + 1 }

    Validate-PeriodMatchesDatabase "CONTATOS.TXT" $LineNumber $ano $trimestre

    if ($tipoContato -eq "I") {
        if ($nomeRaw.Trim().Length -ne 0) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Nome deve ficar em branco quando o tipo de contato é I (institucional)." }
        if ($cargoRaw.Trim().Length -ne 0) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Cargo deve ficar em branco quando o tipo de contato é I (institucional)." }
        if ($telefoneRaw.Trim().Length -ne 0) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Telefone deve ficar em branco quando o tipo de contato é I (institucional)." }
    }
    else {
        Validate-UpperTextFixed "CONTATOS.TXT" $LineNumber "Nome" $nomeRaw 50
        Validate-UpperTextFixed "CONTATOS.TXT" $LineNumber "Cargo" $cargoRaw 50
        if ($telefoneRaw.TrimEnd(' ').Length -eq 0) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Telefone vazio para contato D/T." }
        if ($telefoneRaw.StartsWith(" ")) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Telefone deve ser preenchido da esquerda para a direita, sem espaços à esquerda." }
        if ($telefoneRaw -match '[\x00-\x08\x0B\x0C\x0E-\x1F]') { Add-Error "CONTATOS.TXT linha ${LineNumber}: Telefone contém caractere de controle inválido." }
    }

    if ($email.Length -eq 0) { Add-Error "CONTATOS.TXT linha ${LineNumber}: E-mail vazio." }
    elseif ($email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { Add-Warning "CONTATOS.TXT linha ${LineNumber}: E-mail aparenta estar fora do padrão: '$email'." }
    if ($email -cmatch '[A-Z]') { Add-Error "CONTATOS.TXT linha ${LineNumber}: E-mail deve usar caracteres em minúsculas. Valor útil: '$email'." }
    if ($emailRaw.StartsWith(" ")) { Add-Error "CONTATOS.TXT linha ${LineNumber}: E-mail deve ser preenchido da esquerda para a direita, sem espaços à esquerda." }

    if ($enderecoRaw.TrimEnd(' ').Length -eq 0) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Endereço institucional vazio." }
    if ($enderecoRaw.StartsWith(" ")) { Add-Error "CONTATOS.TXT linha ${LineNumber}: Endereço deve ser preenchido da esquerda para a direita, sem espaços à esquerda." }
}

function Validate-ContatosAggregate {
    if ([int]$script:ContatosCounts["D"] -ne 1) { Add-Error "CONTATOS.TXT: deve haver exatamente 1 contato do tipo D (Diretor). Encontrados: $($script:ContatosCounts["D"])." }
    if ([int]$script:ContatosCounts["T"] -ne 2) { Add-Error "CONTATOS.TXT: deve haver exatamente 2 contatos do tipo T (Técnicos). Encontrados: $($script:ContatosCounts["T"])." }
    if ([int]$script:ContatosCounts["I"] -ne 1) { Add-Error "CONTATOS.TXT: deve haver exatamente 1 contato do tipo I (Institucional). Encontrados: $($script:ContatosCounts["I"])." }
}

function Validate-DescricaDetail {
    param([string]$Line, [int]$LineNumber)

    $ano = $Line.Substring(0, 4); $trimestre = $Line.Substring(4, 1); $proposito = $Line.Substring(5, 1); $modalidade = $Line.Substring(6, 1); $abrangencia = $Line.Substring(7, 1)
    $descricaoRaw = $Line.Substring(8, 1000); $descricao = $descricaoRaw.TrimEnd(' ')

    if (-not (Test-OnlyDigits $ano)) { Add-Error "DESCRICA.TXT linha ${LineNumber}: Ano deve conter exatamente 4 dígitos. Valor: '$ano'." }
    if ($trimestre -notin @("1", "2", "3", "4")) { Add-Error "DESCRICA.TXT linha ${LineNumber}: Trimestre inválido: '$trimestre'." }
    if ($proposito -notin @("1", "2")) { Add-Error "DESCRICA.TXT linha ${LineNumber}: Propósito inválido: '$proposito'." }
    if ($modalidade -notin @("1", "2", "3", "4")) { Add-Error "DESCRICA.TXT linha ${LineNumber}: Modalidade inválida: '$modalidade'." }
    if ($abrangencia -notin @("1", "2")) { Add-Error "DESCRICA.TXT linha ${LineNumber}: Abrangência inválida: '$abrangencia'." }

    Validate-PeriodMatchesDatabase "DESCRICA.TXT" $LineNumber $ano $trimestre

    if ($descricao.Length -eq 0) { Add-Error "DESCRICA.TXT linha ${LineNumber}: descrição está vazia." }
    if ($descricaoRaw.StartsWith(" ")) { Add-Error "DESCRICA.TXT linha ${LineNumber}: descrição deve ser preenchida da esquerda para a direita, sem espaços à esquerda." }
    if ($descricaoRaw -match '[\x00-\x08\x0B\x0C\x0E-\x1F]') { Add-Error "DESCRICA.TXT linha ${LineNumber}: descrição contém caractere de controle inválido." }
    if ($descricao.Length -gt 1000) { Add-Error "DESCRICA.TXT linha ${LineNumber}: descrição possui mais de 1000 caracteres úteis." }

    $key = $Line.Substring(0, 8)
    Add-DuplicateCheck $script:DescricaKeys $key "DESCRICA.TXT" $LineNumber "a mesma combinação de Ano, Trimestre, Propósito, Modalidade e Abrangência"
}

function Validate-Database {
    param([string]$Line)

    if ($Line.Length -ne 30) { Add-Error "DATABASE.TXT: linha única possui $($Line.Length) caracteres; esperado: 30."; return }

    $nomeArquivo = $Line.Substring(0, 8); $dataGeracao = $Line.Substring(8, 8); $iap = $Line.Substring(16, 8); $dataBase = $Line.Substring(24, 6)
    $script:HeaderIapByFile["DATABASE.TXT"] = $iap

    if ($nomeArquivo -ne "DATABASE") { Add-Error "DATABASE.TXT: nome no HEADER inválido. Encontrado '$nomeArquivo'; esperado 'DATABASE'." }
    if (-not (Test-ValidDateYyyyMMdd $dataGeracao)) { Add-Error "DATABASE.TXT: data de geração deve ser uma data válida em AAAAMMDD. Valor: '$dataGeracao'." }
    if (-not (Test-AlphaNum8 $iap)) { Add-Error "DATABASE.TXT: IAP deve conter exatamente 8 caracteres alfanuméricos, sem máscara/pontuação/espaços. Valor: '$iap'." }

    if (-not (Test-OnlyDigits $dataBase)) { Add-Error "DATABASE.TXT: data-base deve ser numérica AAAAMM. Valor: '$dataBase'."; return }
    $mes = $dataBase.Substring(4, 2)
    if ($mes -notin @("03", "06", "09", "12")) { Add-Error "DATABASE.TXT: data-base inválida '$dataBase'. O mês deve ser 03, 06, 09 ou 12." }
    if ([int]$dataBase -lt 201812) { Add-Error "DATABASE.TXT: data-base '$dataBase' é anterior ao primeiro período válido (201812)." }

    $script:DatabaseDataBase = $dataBase
    $script:DatabaseAno = $dataBase.Substring(0, 4)
    $script:DatabaseTrimestre = Get-QuarterFromMonth $mes
}

function Validate-HeaderIapConsistency {
    if ($script:HeaderIapByFile.Count -eq 0) { return }
    $iapRef = $script:HeaderIapByFile["DATABASE.TXT"]
    if (-not $iapRef) { return }
    foreach ($file in $script:HeaderIapByFile.Keys) {
        if ($script:HeaderIapByFile[$file] -ne $iapRef) { Add-Error "$file: IAP do HEADER ('$($script:HeaderIapByFile[$file])') diverge do IAP do DATABASE.TXT ('$iapRef')." }
    }
}

if (-not (Test-Path $ZipPath)) { throw "Arquivo não encontrado: $ZipPath" }

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
            if ($entry.FullName.EndsWith("/")) { Add-Error "O ZIP contém pasta/diretório: '$($entry.FullName)'. Os arquivos devem estar na raiz."; continue }
            if ($entry.FullName -ne $entry.Name) { Add-Error "O arquivo '$($entry.FullName)' está dentro de uma pasta. Os arquivos devem estar na raiz do ZIP." }
        }

        $actualFiles = $entries | Where-Object { -not $_.FullName.EndsWith("/") } | ForEach-Object { $_.Name.ToUpperInvariant() }

        foreach ($expected in $ExpectedFiles) { if ($actualFiles -notcontains $expected) { Add-Error "Arquivo obrigatório ausente no ZIP: $expected" } }
        foreach ($actual in $actualFiles) { if ($ExpectedFiles -notcontains $actual) { Add-Error "Arquivo extra ou nome incorreto encontrado no ZIP: $actual" } }

        foreach ($entry in $entries) {
            if ($entry.FullName.EndsWith("/")) { continue }
            $target = Join-Path $tempDir $entry.Name
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, $true)
        }
    }
    finally { $zip.Dispose() }

    # DATABASE.TXT é processado primeiro para permitir validação de período nos demais arquivos.
    $orderedFiles = @("DATABASE.TXT") + ($ExpectedFiles | Where-Object { $_ -ne "DATABASE.TXT" })

    foreach ($fileName in $orderedFiles) {
        $filePath = Join-Path $tempDir $fileName
        if (-not (Test-Path $filePath)) { continue }

        $content = Get-TextIso88591 $filePath
        $lines = Split-LinesStrict $content
        if ($lines.Count -eq 0) { Add-Error "$fileName: arquivo vazio."; continue }

        Validate-NoSeparators $fileName $lines
        $spec = $Layout[$fileName]

        if ($fileName -eq "DATABASE.TXT") {
            if ($lines.Count -ne 1) { Add-Error "DATABASE.TXT deve possuir somente uma linha. Encontradas: $($lines.Count)." }
            Validate-Database $lines[0]
            continue
        }

        $header = $lines[0]
        $details = @()
        if ($lines.Count -gt 1) { $details = $lines[1..($lines.Count - 1)] }

        Validate-HeaderCommon $fileName $header $spec
        Validate-RecordCount $fileName $header $details.Count

        for ($i = 0; $i -lt $details.Count; $i++) {
            $line = $details[$i]
            $lineNumber = $i + 2
            if ($line.Length -ne $spec.DetailLength) { Add-Error "$fileName linha ${lineNumber}: possui $($line.Length) caracteres; esperado: $($spec.DetailLength)."; continue }
            switch ($fileName) {
                "TARIFAS.TXT"  { Validate-TarifasDetail $line $lineNumber }
                "PARTICIP.TXT" { Validate-ParticipDetail $line $lineNumber }
                "CONTATOS.TXT" { Validate-ContatosDetail $line $lineNumber }
                "DESCRICA.TXT" { Validate-DescricaDetail $line $lineNumber }
            }
        }
    }

    Validate-HeaderIapConsistency
    if (Test-Path (Join-Path $tempDir "CONTATOS.TXT")) { Validate-ContatosAggregate }

    Write-Host ""
    Write-Host "=============================="
    Write-Host "VALIDAÇÃO CADOC 6333"
    Write-Host "=============================="
    Write-Host ""

    if ($Warnings.Count -gt 0) {
        Write-Host "AVISOS:" -ForegroundColor Yellow
        foreach ($warning in $Warnings) { Write-Host " - $warning" -ForegroundColor Yellow }
        Write-Host ""
    }

    if ($Errors.Count -gt 0) {
        Write-Host "ERROS:" -ForegroundColor Red
        foreach ($errorItem in $Errors) { Write-Host " - $errorItem" -ForegroundColor Red }
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
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
}
