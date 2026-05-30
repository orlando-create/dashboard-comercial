# Setup: Deploy automático no Netlify

## Pré-requisitos (fazer uma vez)

### 1. Instalar Node.js
Baixe em https://nodejs.org (versão LTS)

### 2. Instalar Netlify CLI
```
npm install -g netlify-cli
```

### 3. Fazer login e linkar o site
```
cd "C:\Users\orlan\OneDrive\Documentos\Claude\Projects\Indicadores Time Comercial"
netlify login
netlify link
```
> Quando perguntar o site, escolha **indicadorescomercial** (pelo nome ou URL)

---

## Criar tarefa no Windows Task Scheduler (8h e 18h)

### Abra o PowerShell como administrador e rode:

```powershell
# Tarefa das 8h
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -File `"C:\Users\orlan\OneDrive\Documentos\Claude\Projects\Indicadores Time Comercial\deploy_local.ps1`""

$trigger8 = New-ScheduledTaskTrigger -Daily -At "08:05AM"
$trigger18 = New-ScheduledTaskTrigger -Daily -At "06:05PM"

Register-ScheduledTask `
    -TaskName "Dashboard Comercial - Atualizar e Publicar" `
    -Action $action `
    -Trigger $trigger8, $trigger18 `
    -RunLevel Highest `
    -Force
```

---

## Rodar manualmente
```
powershell -ExecutionPolicy Bypass -File deploy_local.ps1
```

---

## Como funciona
1. `regenerate_dashboard.py` lê o Excel e atualiza o `One_Page_Comercial_2026.html`
2. `netlify deploy --prod` publica o HTML no ar em ~30 segundos
3. Log em: `automation/last_run.log`
