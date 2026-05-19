#!/usr/bin/env python3
"""
Patch pontual: insere a coluna 'Discricionário' na aba Base do Sheets,
entre 'Observações' e 'Origem', com a mesma formatação azul dos demais cabeçalhos.
Execute uma única vez; execuções seguintes detectam que já existe e encerram.
"""
import os, json, sys
from google.oauth2.service_account import Credentials
import gspread

GSHEET_CREDS = os.environ.get("GSHEET_CREDENTIALS", "")
GSHEET_ID    = os.environ.get("GSHEET_ID", "")

if not GSHEET_CREDS or not GSHEET_ID:
    print("ERRO: defina GSHEET_CREDENTIALS e GSHEET_ID antes de rodar.")
    sys.exit(1)

creds = Credentials.from_service_account_info(
    json.loads(GSHEET_CREDS),
    scopes=[
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]
)
gc  = gspread.authorize(creds)
sh  = gc.open_by_key(GSHEET_ID)
aba = sh.worksheet('Base')
sid = aba.id

# ── Lê cabeçalhos atuais (linha 2) ──────────────────────────────────────────
headers = aba.row_values(2)
print(f"Colunas atuais ({len(headers)}): ...{headers[-4:]}")

if 'Discricionário' in headers:
    print("Coluna 'Discricionário' já existe — nada a fazer.")
    sys.exit(0)

# Insere ANTES de 'Origem' (ou no final se não encontrar)
try:
    idx_origem = headers.index('Origem')   # 0-based
except ValueError:
    idx_origem = len(headers)

col_insert  = idx_origem + 1               # 1-based (para update_cell)
n_cols_novo = len(headers) + 1             # total após inserção

def _rgb(r, g, b):
    return {'red': round(r/255, 4), 'green': round(g/255, 4), 'blue': round(b/255, 4)}

C_AZUL   = _rgb(0,   48,  135)   # #003087
C_BRANCO = _rgb(255, 255, 255)

# ── batch: inserir coluna + corrigir merge do título + formatar cabeçalho ────
sh.batch_update({'requests': [

    # 1. Insere coluna vazia em idx_origem (empurra 'Origem' uma posição à direita)
    {'insertDimension': {
        'range': {
            'sheetId':   sid,
            'dimension': 'COLUMNS',
            'startIndex': idx_origem,
            'endIndex':   idx_origem + 1,
        },
        'inheritFromBefore': True,
    }},

    # 2. Desfaz merge atual do título (range pré-inserção pode ter ficado defasado)
    {'unmergeCells': {
        'range': {
            'sheetId':          sid,
            'startRowIndex':    0, 'endRowIndex':    1,
            'startColumnIndex': 0, 'endColumnIndex': n_cols_novo,
        },
    }},

    # 3. Refaz merge do título cobrindo todas as colunas (incluindo a nova)
    {'mergeCells': {
        'range': {
            'sheetId':          sid,
            'startRowIndex':    0, 'endRowIndex':    1,
            'startColumnIndex': 0, 'endColumnIndex': n_cols_novo,
        },
        'mergeType': 'MERGE_ALL',
    }},

    # 4. Aplica formato azul no novo cabeçalho (linha 2, posição idx_origem)
    {'repeatCell': {
        'range': {
            'sheetId':          sid,
            'startRowIndex':    1, 'endRowIndex':    2,
            'startColumnIndex': idx_origem,
            'endColumnIndex':   idx_origem + 1,
        },
        'cell': {'userEnteredFormat': {
            'backgroundColor':     C_AZUL,
            'horizontalAlignment': 'CENTER',
            'verticalAlignment':   'MIDDLE',
            'wrapStrategy':        'WRAP',
            'textFormat': {
                'foregroundColor': C_BRANCO,
                'bold':            True,
                'fontFamily':      'Arial',
                'fontSize':        11,
            },
        }},
        'fields': ('userEnteredFormat(backgroundColor,horizontalAlignment,'
                   'verticalAlignment,wrapStrategy,textFormat)'),
    }},
]})

# Escreve o texto do cabeçalho (após o batchUpdate para garantir que a célula existe)
aba.update_cell(2, col_insert, 'Discricionário')

print(f"✓ 'Discricionário' inserida na coluna {col_insert} (entre Observações e Origem).")
print(f"  Título remergeado sobre {n_cols_novo} colunas.")
print("  Dados da coluna serão preenchidos automaticamente no próximo run do Actions.")
