#!/usr/bin/env python3
"""
Baixa os 5 arquivos-fonte do Google Drive (master, Controle Vendas, e as 3
planilhas de vendedor) como .xlsx, usando uma service account do Google.
Isso substitui o passo manual de "exportar e anexar" — depois de feito o
setup abaixo (uma vez só), rodar este script sempre deixa os 5 arquivos
atualizados na pasta de trabalho, prontos pra build_dashboard.py.

AUTENTICAÇÃO EM USO (desde 15/10/2026): Workload Identity Federation, SEM
chave JSON. A org do Google bloqueia criação de chaves
(iam.disableServiceAccountKeyCreation), então o GitHub Actions do repo
psalesc/phibo-dashboard entra como a service account
dashboard-reader@phibo-dashboard-511718.iam.gserviceaccount.com
(pool "github", provider "phibo-dashboard", projeto phibo-dashboard-511718).
Sem service_account.json presente, o script usa google.auth.default().
O passo a passo abaixo (chave JSON) fica só como referência.

SETUP ANTIGO COM CHAVE (bloqueado pela org — referência):
1. Acesse console.cloud.google.com e crie um projeto (ou use um existente).
2. Nesse projeto, ative a "Google Drive API" (APIs e serviços > Biblioteca).
3. Crie uma Service Account (IAM e Admin > Contas de serviço > Criar conta
   de serviço). Não precisa de nenhuma permissão/role especial no projeto.
4. Nessa service account, crie uma chave: aba "Chaves" > Adicionar chave >
   Criar nova chave > JSON. Isso baixa um arquivo .json.
5. Salve esse arquivo como "service_account.json" na MESMA pasta deste
   script (pipeline/). NUNCA faça commit desse arquivo no git — ele já
   está listado no .gitignore deste repositório.
6. Abra o arquivo JSON e copie o valor do campo "client_email" (algo como
   nome@projeto.iam.gserviceaccount.com).
7. No Google Drive, compartilhe CADA UMA das 5 planilhas abaixo com esse
   e-mail, como Leitor (Viewer) — é só clicar em "Compartilhar" em cada
   uma e colar o e-mail, igual compartilhar com uma pessoa.

Depois do setup, uso:
    python3 fetch_sheets.py [pasta_saida]

Sem argumento, salva na pasta atual. Os nomes de saída (master_latest.xlsx,
controle_vendas.xlsx, samuel.xlsx, marina.xlsx, bruno.xlsx) já são os que
build_dashboard.py / build_historical.py esperam.

Dependências (pip): google-api-python-client, google-auth
    pip install google-api-python-client google-auth
"""
import sys
import os
import io

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_PATH = os.path.join(SCRIPT_DIR, 'service_account.json')

# IDs dos arquivos no Google Drive (extraídos das URLs originais — ver
# CLAUDE.md na raiz do repositório pra contexto de cada planilha). Esses
# IDs são estáveis mesmo que o Pedro renomeie o arquivo ou mude o conteúdo.
FILES = {
    'master_latest.xlsx': '1YFK6s_fJw_VIVeQWJ-p9qtqtY39L2Qakx7jN6WvXCXg',   # Dashboard de Vendas — Phibo (ao vivo)
    'controle_vendas.xlsx': '15jgTNMrhClhy4h7c-rURpVg6IfJd4uPAU9Qe8_k2GoE',  # Controle Vendas
    'samuel.xlsx': '1z78YkW9VIvSLEnJ2iTGI9CEXNeMYbXDHPnlxJk6MizA',           # Samuel - Planilha
    'marina.xlsx': '1Y_tPhMDUiYJD6zRlnZLIEvYvkVPNlCTyyaH2DMgH1H4',           # Marina - Planilha
    'bruno.xlsx': '1JYUSwj97Y2huGglji619t1UkoXaKCoOoUbsIzRYPTEQ',            # Bruno - Planilha
}

XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']


def main():
    out_dir = sys.argv[1] if len(sys.argv) > 1 else '.'
    os.makedirs(out_dir, exist_ok=True)

    try:
        import google.auth
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload
    except ImportError:
        print('ERRO: dependências não instaladas. Rode:')
        print('    pip install google-api-python-client google-auth')
        sys.exit(1)

    if os.path.exists(CREDENTIALS_PATH):
        creds = service_account.Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
    else:
        # Modo sem chave (o que está em uso): no GitHub Actions, o passo
        # google-github-actions/auth troca o token OIDC do GitHub por uma
        # credencial da service account via Workload Identity Federation e
        # aponta GOOGLE_APPLICATION_CREDENTIALS pra ela.
        try:
            creds, _ = google.auth.default(scopes=SCOPES)
        except google.auth.exceptions.DefaultCredentialsError:
            print('ERRO: nenhuma credencial do Google encontrada.')
            print('Este script roda pelo GitHub Actions (Workload Identity Federation,')
            print('sem chave). Pra rodar fora dele, baixe os 5 arquivos manualmente do')
            print('Drive (Arquivo > Fazer download > Microsoft Excel) com os mesmos nomes.')
            sys.exit(1)
    service = build('drive', 'v3', credentials=creds)

    ok, failed = 0, []
    for out_name, file_id in FILES.items():
        out_path = os.path.join(out_dir, out_name)
        try:
            request = service.files().export_media(fileId=file_id, mimeType=XLSX_MIME)
            buf = io.BytesIO()
            downloader = MediaIoBaseDownload(buf, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            with open(out_path, 'wb') as f:
                f.write(buf.getvalue())
            print(f'OK: {out_name} ({len(buf.getvalue())} bytes)')
            ok += 1
        except Exception as e:
            print(f'FALHOU: {out_name} ({file_id}) — {e}')
            failed.append(out_name)

    if failed:
        print(f'\n{len(failed)} arquivo(s) falharam: {failed}')
        print('Causa mais comum: a planilha não foi compartilhada com o e-mail')
        print('da service account (client_email no JSON) — ou foi compartilhada')
        print('depois de o arquivo já estar em cache. Confira o compartilhamento')
        print('de cada planilha no Drive.')
        sys.exit(1)
    print(f'\n{ok}/{len(FILES)} arquivos baixados com sucesso em "{out_dir}".')


if __name__ == '__main__':
    main()
