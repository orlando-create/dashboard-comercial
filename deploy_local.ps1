# ============================================================
# Deploy One Page Comercial 2026 -> Netlify
# ============================================================

$BASE = Split-Path -Parent $MyInvocation.MyCommand.Path
$LOG  = "$BASE\automation\last_run.log"

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $msg"
    Write-Host $line
    Add-Content -Path $LOG -Value $line -Encoding UTF8
}

Log "============================================================"
Log "Iniciando deploy no Netlify..."

# Verifica se o HTML existe
$HTML = "$BASE\One_Page_Comercial_2026.html"
if (-not (Test-Path $HTML)) {
    Log "[ERRO] Arquivo nao encontrado: $HTML"
    exit 1
}
Log "HTML encontrado: $HTML"

# Verifica se netlify-cli esta instalado
$netlifyCmd = Get-Command netlify -ErrorAction SilentlyContinue
if (-not $netlifyCmd) {
    Log "[AVISO] netlify-cli nao encontrado. Instale com: npm install -g netlify-cli"
    exit 1
}

# Cria pasta temporaria com o HTML como index.html
$DEPLOY_DIR = "$BASE\automation\_deploy_tmp"
if (Test-Path $DEPLOY_DIR) { Remove-Item $DEPLOY_DIR -Recurse -Force }
New-Item -ItemType Directory -Path $DEPLOY_DIR | Out-Null
Copy-Item $HTML "$DEPLOY_DIR\index.html"
Log "index.html copiado para pasta de deploy."

# Faz deploy via cmd para evitar problemas com script .ps1 do netlify-cli
Log "Executando netlify deploy..."
cmd /c "cd /d `"$BASE`" && netlify deploy --prod --dir `"$DEPLOY_DIR`""
$exitCode = $LASTEXITCODE

# Limpa pasta temporaria
Remove-Item $DEPLOY_DIR -Recurse -Force -ErrorAction SilentlyContinue

if ($exitCode -eq 0) {
    Log "[OK] Deploy publicado: https://indicadorescomercial.netlify.app"
} else {
    Log "[ERRO] Deploy falhou (codigo $exitCode)"
    exit 1
}

Log "Fim - SUCESSO"
