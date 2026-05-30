#!/usr/bin/env python3
"""
Faz git commit + push do HTML gerado para o GitHub.
O GitHub Actions cuida do deploy no Netlify automaticamente.

Uso: python git_push.py
Requer: github_config.json com { "token": "ghp_...", "repo": "usuario/repo", "branch": "main" }
"""

import subprocess, json, os, datetime, sys

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, 'github_config.json')
HTML_FILE   = 'One_Page_Comercial_2026.html'
LOG_PATH    = os.path.join(BASE_DIR, 'automation', 'last_run.log')


def log(msg):
    ts   = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f'[{ts}] {msg}'
    print(line, flush=True)
    try:
        with open(LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception:
        pass


def run(cmd, cwd=None):
    result = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"Comando falhou: {' '.join(cmd)}\nSTDERR: {result.stderr.strip()}")
    return result.stdout.strip()


def main():
    if not os.path.exists(CONFIG_PATH):
        log('ERRO: github_config.json não encontrado.')
        log('Crie o arquivo com: {"token":"ghp_SEU_TOKEN","repo":"usuario/repo","branch":"main"}')
        sys.exit(1)

    with open(CONFIG_PATH, encoding='utf-8') as f:
        cfg = json.load(f)

    token      = cfg['token']
    repo       = cfg['repo']   # ex: "orlando-strategicos/dashboard-comercial"
    branch     = cfg.get('branch', 'main')
    remote_url = f'https://x-token:{token}@github.com/{repo}.git'

    log(f'[GitHub] Iniciando git push → {repo} ({branch})')

    try:
        # Configura identidade git
        run(['git', 'config', 'user.email', 'claude-automation@strategicos.com.br'], cwd=BASE_DIR)
        run(['git', 'config', 'user.name', 'Claude Automation'], cwd=BASE_DIR)

        # Atualiza remote com token
        remotes = run(['git', 'remote'], cwd=BASE_DIR)
        if 'origin' in remotes.split():
            run(['git', 'remote', 'set-url', 'origin', remote_url], cwd=BASE_DIR)
        else:
            run(['git', 'remote', 'add', 'origin', remote_url], cwd=BASE_DIR)

        # Adiciona o HTML
        run(['git', 'add', HTML_FILE], cwd=BASE_DIR)

        # Verifica se há mudança staged
        diff_check = subprocess.run(
            ['git', 'diff', '--cached', '--quiet'],
            cwd=BASE_DIR
        )
        if diff_check.returncode == 0:
            log('[GitHub] Nenhuma mudança no HTML — push desnecessário.')
            return

        # Commit
        ts      = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
        msg     = f'Auto-update dashboard {ts}'
        run(['git', 'commit', '-m', msg], cwd=BASE_DIR)
        log(f'[GitHub] Commit criado: {msg}')

        # Push
        run(['git', 'push', 'origin', branch], cwd=BASE_DIR)
        log(f'[GitHub] ✅ Push realizado → github.com/{repo}')
        log('[GitHub] GitHub Actions iniciará o deploy no Netlify automaticamente (~1 min).')

    except RuntimeError as e:
        log(f'[GitHub] ERRO: {e}')
        sys.exit(1)


if __name__ == '__main__':
    main()
