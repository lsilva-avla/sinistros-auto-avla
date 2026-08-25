#!/usr/bin/env python3
"""
Backfill de Sinistros — AVLA Crédito
Coleta todos os registros de 2026 usando AMBOS os remetentes históricos:
  - notificaciones-01@avla.com  (jan–jul/2026)
  - info-31@avla.com            (ago/2026 em diante)
Reconstrói a base no Google Sheets e envia email com resumo.
Uso: executar manualmente via GitHub Actions ("Run workflow").
"""

import imaplib
import email
import smtplib
import re
import os
import sys
import json
import calendar
import urllib.request
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
from datetime import datetime, timedelta, date
from bs4 import BeautifulSoup
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# =============================================================
# CONFIGURAÇÃO
# =============================================================
EMAIL_CAIXA   = os.environ.get("EMAIL_CAIXA",   "mgignon@avla.com")
APP_PASSWORD  = os.environ.get("APP_PASSWORD",   "")
_destino_env  = os.environ.get("EMAIL_DESTINO",  "")
DESTINATARIOS = [e.strip() for e in _destino_env.split(',') if e.strip()] \
                if _destino_env else ["lsilva@avla.com", "mgignon@avla.com"]
GSHEET_CREDS  = os.environ.get("GSHEET_CREDENTIALS", "")
GSHEET_ID     = os.environ.get("GSHEET_ID", "")

# Remetentes a considerar (ordem de prioridade irrelevante — todos são válidos)
REMETENTES_VALIDOS = [
    "notificaciones-01@avla.com",   # sender antigo (jan–jul/2026)
    "info-31@avla.com",             # sender novo   (ago/2026+)
]

IMAP_SERVER = "imap.gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT   = 587

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if not APP_PASSWORD:
    print("ERRO: APP_PASSWORD não definido.")
    sys.exit(1)
# =============================================================


# ── IMAP helpers ────────────────────────────────────────────

def _select(mail, nome):
    try:
        nome_imap = f'"{nome}"' if any(c in nome for c in '[] ') else nome
        typ, data = mail.select(nome_imap)
        if typ == 'OK':
            count = data[0].decode() if data and data[0] else '?'
            return True, count
        return False, 0
    except Exception:
        return False, 0


def selecionar_pasta_completa(mail):
    try:
        _, pastas = mail.list()
        for item in pastas:
            if item is None:
                continue
            decoded = item.decode('utf-8') if isinstance(item, bytes) else str(item)
            if '\\All' in decoded:
                m = re.search(r'"([^"]+)"\s*$|(\S+)\s*$', decoded)
                if m:
                    nome = (m.group(1) or m.group(2)).strip()
                    ok, count = _select(mail, nome)
                    if ok:
                        print(f"✓ Pasta: {nome} — {count} msgs")
                        return
    except Exception as e:
        print(f"Erro ao detectar All Mail: {e}")
    for nome in ['[Gmail]/All Mail', '[Gmail]/Todos os e-mails', 'INBOX']:
        ok, count = _select(mail, nome)
        if ok:
            print(f"✓ Pasta (fallback): {nome} — {count} msgs")
            return
    raise RuntimeError("Não foi possível selecionar pasta de email.")


# ── Coleta com múltiplos remetentes ─────────────────────────

def _buscar_com_remetente(mail, remetente, since, before, vistos):
    """Busca emails de SINISTRO de um remetente específico no período."""
    try:
        _, ids = mail.search(
            None,
            f'(FROM "{remetente}" SUBJECT "SINISTRO" SINCE "{since}" BEFORE "{before}")'
        )
    except Exception as e:
        print(f"  Erro IMAP ({remetente}): {e}")
        return [], vistos

    ids_list = ids[0].split() if ids and ids[0] else []
    print(f"  FROM {remetente}: {len(ids_list)} email(s) encontrado(s)")

    novos = []
    for msg_id in ids_list:
        try:
            _, dados = mail.fetch(msg_id, '(RFC822)')
            msg = email.message_from_bytes(dados[0][1])
        except Exception:
            continue

        mid = msg.get('Message-ID', f'_noid_{msg_id.decode()}_')
        if mid in vistos:
            continue
        vistos.add(mid)

        assunto    = str(msg.get('Subject', ''))
        data_email = email.utils.parsedate_to_datetime(msg['Date'])

        html_body = None
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                payload   = part.get_payload(decode=True)
                charset   = part.get_content_charset() or 'utf-8'
                html_body = payload.decode(charset, errors='ignore')
                break

        if not html_body:
            continue

        campos      = parse_corpo_email(html_body)
        num_assunto = extrair_numero_sinistro(assunto)
        num_corpo   = campos.get('N° Sinistro', '') or _extrair_num_corpo(html_body)
        num_final   = num_assunto or num_corpo
        if not num_final:
            print(f"    ⚠ N° Sinistro não encontrado | {assunto[:60]}")
        campos['N° Sinistro'] = num_final
        campos['Data']        = data_email.strftime('%d/%m/%Y')
        novos.append(campos)

    return novos, vistos


def coletar_emails_backfill(inicio: datetime, fim: datetime) -> list:
    print(f"Conectando em {EMAIL_CAIXA}...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_CAIXA, APP_PASSWORD)

    since  = inicio.strftime("%d-%b-%Y")
    before = (fim + timedelta(days=1)).strftime("%d-%b-%Y")
    print(f"Período IMAP: {since} → {before}")

    selecionar_pasta_completa(mail)

    vistos    = set()
    registros = []

    for remetente in REMETENTES_VALIDOS:
        novos, vistos = _buscar_com_remetente(mail, remetente, since, before, vistos)
        registros.extend(novos)

    mail.logout()
    print(f"Total coletado (deduplicado): {len(registros)} sinistros")
    return registros


# ── Parse ────────────────────────────────────────────────────

def extrair_numero_sinistro(assunto: str) -> str:
    padroes = [
        r'N[°º]?\s*(?:DE\s+)?S[Ii][Nn][Ii][Ee]?[Ss][Tt][Rr][Oo]\s*[:\-]?\s*(\d+)',
        r'S[Ii][Nn][Ii][Ee]?[Ss][Tt][Rr][Oo]\s*[N°nº#]?\s*[:\-]?\s*(\d+)',
        r'\b(\d{6,})\b',
    ]
    for p in padroes:
        m = re.search(p, assunto, re.IGNORECASE)
        if m:
            return m.group(1)
    return ''


def _extrair_num_corpo(html: str) -> str:
    soup = BeautifulSoup(html, 'html.parser')
    texto = soup.get_text(separator=' ')
    padroes = [
        r'N[°º\.]\s*(?:de\s+)?[Ss]ini[es]stro\s*[:\-]?\s*(\d{4,})',
        r'[Ss]ini[es]stro\s*[N°nº#\.]*\s*[:\-]?\s*(\d{4,})',
        r'(?:N[°º]|No\.?|Nro\.?|Nr\.?)\s*[:\-]?\s*(\d{5,})',
        r'\b(\d{8,})\b',
    ]
    for p in padroes:
        m = re.search(p, texto, re.IGNORECASE)
        if m:
            return m.group(1)
    return ''


def parse_corpo_email(html: str) -> dict:
    soup = BeautifulSoup(html, 'html.parser')
    texto = soup.get_text(separator='\n')

    campos_aliases = {
        'N° Sinistro':      ['N° Sinistro', 'Nº Sinistro', 'N° de Sinistro',
                             'Número do Sinistro', 'Número de Sinistro',
                             'N° de Siniestro', 'Nº de Siniestro',
                             'Número de Siniestro', 'No. de Siniestro',
                             'No de Siniestro', 'Siniestro'],
        'Segurado':         ['Segurado'],
        'Filial':           ['Filial'],
        'Apólice':          ['Apólice', 'Apolice', 'Póliza', 'Poliza'],
        'Devedor':          ['Devedor', 'Deudor'],
        'CNPJ do Devedor':  ['CNPJ do devedor', 'CNPJ do Devedor', 'CNPJ',
                             'RUT', 'RUT del Deudor'],
        'Ocorrência':       ['Ocorrência', 'Ocorrencia', 'Siniestro',
                             'Tipo de Siniestro'],
        'Declaração':       ['Declaração', 'Declaracao', 'Fecha de Declaración',
                             'Fecha Declaracion'],
        'Valor Sinistrado': ['Valor sinistrado', 'Valor Sinistrado',
                             'Monto Sinistrado', 'Monto del Siniestro',
                             'Valor do Sinistro'],
    }

    resultado = {}
    for tag in soup.find_all(['b', 'strong']):
        label = tag.get_text(strip=True).rstrip(':')
        for campo, aliases in campos_aliases.items():
            if campo in resultado:
                continue
            for alias in aliases:
                if alias.lower() == label.lower():
                    proximo = tag.next_sibling
                    if proximo:
                        if hasattr(proximo, 'get_text'):
                            valor = proximo.get_text(strip=True).lstrip(':').strip()
                        else:
                            valor = str(proximo).strip().lstrip(':').strip()
                        if valor:
                            resultado[campo] = valor
                    break

    for campo, aliases in campos_aliases.items():
        if campo in resultado:
            continue
        for alias in aliases:
            padrao = re.compile(rf'{re.escape(alias)}\s*:?\s*(.+)', re.IGNORECASE)
            m = padrao.search(texto)
            if m:
                resultado[campo] = m.group(1).strip()
                break

    return resultado


# ── Enriquecimento e Sheets (reutilizado de relatorio_diario.py) ──

def _tc(s) -> str:
    v = str(s).strip() if s is not None else ''
    if not v or v.lower() in ('nan', 'none', ''):
        return ''
    return v.title()


def _normalizar_apolice(v) -> str:
    s = str(v).strip() if v is not None else ''
    if not s or s.lower() in ('nan', 'none', ''):
        return ''
    try:
        return str(int(float(s)))
    except (ValueError, TypeError):
        return re.sub(r'\D', '', s)


def _normalizar_cnpj(v) -> str:
    digits = re.sub(r'\D', '', str(v or ''))
    if not digits:
        return ''
    if len(digits) == 13:
        digits = digits.zfill(14)
    elif len(digits) == 10:
        digits = digits.zfill(11)
    return digits


def _parse_date(s: str) -> date:
    try:
        return datetime.strptime(s, '%d/%m/%Y').date()
    except (ValueError, TypeError):
        return date(2000, 1, 1)


def _key_decl(r: dict) -> date:
    try:
        return datetime.strptime(r.get('Declaração', '') or '', '%d/%m/%Y').date()
    except (ValueError, TypeError):
        return date(2000, 1, 1)


_ptax_cache: dict = {}


def _get_ptax(ano: int, mes: int, moeda: str = 'USD') -> float:
    chave = (moeda, ano, mes)
    if chave in _ptax_cache:
        return _ptax_cache[chave]
    hoje   = date.today()
    ultimo = calendar.monthrange(ano, mes)[1]
    if ano == hoje.year and mes == hoje.month:
        ultimo = min(ultimo, (hoje - timedelta(days=1)).day)
    for delta in range(0, 10):
        d = date(ano, mes, ultimo) - timedelta(days=delta)
        if d.month != mes:
            break
        data_str = d.strftime('%m-%d-%Y')
        url = (
            'https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/'
            f'CotacaoDolarDia(dataCotacao=@d)?@d=%27{data_str}%27&$format=json&$select=cotacaoVenda'
            if moeda == 'USD' else
            'https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/'
            f'CotacaoMoedaDia(moeda=@m,dataCotacao=@d)?@m=%27{moeda}%27&@d=%27{data_str}%27&$format=json&$select=cotacaoVenda'
        )
        try:
            req = urllib.request.Request(url, headers={'Accept': 'application/json'})
            with urllib.request.urlopen(req, timeout=8) as resp:
                valores = json.loads(resp.read()).get('value', [])
                if valores:
                    taxa = float(valores[-1]['cotacaoVenda'])
                    _ptax_cache[chave] = taxa
                    return taxa
        except Exception:
            pass
    _ptax_cache[chave] = 0.0
    return 0.0


def _detect_moeda(s: str):
    s = str(s).strip() if s else ''
    if not s or s.lower() in ('nan', 'none', ''):
        return 'BRL', 0.0
    s_upper = s.upper()
    if any(tok in s_upper for tok in ('US$', 'USS', 'USD', 'U$S')):
        moeda = 'USD'
    elif '€' in s or 'EUR' in s_upper:
        moeda = 'EUR'
    else:
        moeda = 'BRL'
    nums = re.sub(r'[^\d.,]', '', s)
    if not nums:
        return moeda, 0.0
    if ',' in nums:
        nums = nums.replace('.', '').replace(',', '.')
    elif nums.count('.') > 1:
        nums = nums.replace('.', '')
    elif '.' in nums:
        after_dot = nums.split('.')[1]
        if len(after_dot) == 3:
            nums = nums.replace('.', '')
    try:
        return moeda, float(nums)
    except (ValueError, TypeError):
        return moeda, 0.0


def enriquecer(registros: list) -> list:
    path_cnae  = os.path.join(BASE_DIR, 'dados', 'SEGURADOS_CNAE.xlsx')
    path_grupo = os.path.join(BASE_DIR, 'dados', 'GRUPO_ECONOMICO.xlsx')
    lookup_cnae  = {}
    lookup_grupo = {}

    if os.path.exists(path_cnae):
        try:
            df = pd.read_excel(path_cnae)
            cols = {str(c).lower(): c for c in df.columns}
            col_pol = next((cols[c] for c in cols if 'poliz' in c or 'apolic' in c), df.columns[0])
            col_vi  = next((cols[c] for c in cols if 'inicio' in c), None)
            col_vf  = next((cols[c] for c in cols if 'fim' in c), None)
            col_set = next((cols[c] for c in cols if 'setor' == c), None)
            col_sub = next((cols[c] for c in cols if 'subsetor' == c), None)
            df[col_pol] = df[col_pol].apply(_normalizar_apolice)
            for _, row in df.iterrows():
                chave = row[col_pol]
                if not chave:
                    continue
                vi = row.get(col_vi, '') if col_vi else ''
                vf = row.get(col_vf, '') if col_vf else ''
                lookup_cnae[chave] = {
                    'Vigencia Inicio': vi.strftime('%d/%m/%Y') if (hasattr(vi, 'strftime') and vi.year >= 1950) else '',
                    'Vigencia Fim':    vf.strftime('%d/%m/%Y') if (hasattr(vf, 'strftime') and vf.year >= 1950) else '',
                    'SETOR':    _tc(row.get(col_set, '')) if col_set else '',
                    'SUBSETOR': _tc(row.get(col_sub, '')) if col_sub else '',
                }
            print(f"CNAE: {len(lookup_cnae)} apólices.")
        except Exception as e:
            print(f"Erro CNAE: {e}")

    if os.path.exists(path_grupo):
        try:
            df = pd.read_excel(path_grupo)
            cols = {str(c).lower(): c for c in df.columns}
            col_id    = next((cols[c] for c in cols if 'identificador' in c), df.columns[2])
            col_grupo = next((cols[c] for c in cols if 'grupo' in c), df.columns[4])
            for _, row in df.iterrows():
                d = _normalizar_cnpj(row[col_id])
                g = _tc(row[col_grupo])
                if d and g and d not in lookup_grupo:
                    lookup_grupo[d] = g
            print(f"Grupo Econômico: {len(lookup_grupo)} participantes.")
        except Exception as e:
            print(f"Erro Grupo: {e}")

    for r in registros:
        for campo in ('Segurado', 'Filial', 'Devedor', 'Ocorrência', 'Declaração'):
            if r.get(campo):
                r[campo] = _tc(r[campo])

        apolice   = _normalizar_apolice(r.get('Apólice', ''))
        cnae_info = lookup_cnae.get(apolice, {})
        r['Vigencia Inicio'] = cnae_info.get('Vigencia Inicio', '')
        r['Vigencia Fim']    = cnae_info.get('Vigencia Fim', '')
        r['SETOR']           = cnae_info.get('SETOR', '')
        r['SUBSETOR']        = cnae_info.get('SUBSETOR', '')

        cnpj_digits          = _normalizar_cnpj(r.get('CNPJ do Devedor', ''))
        r['Grupo Econômico'] = lookup_grupo.get(cnpj_digits, '')

        d = _parse_date(r.get('Data', ''))
        r['Data_ISO'] = d.strftime('%Y-%m-%d') if d.year >= 2020 else ''
        r['Mes_Ano']  = d.strftime('%Y-%m')    if d.year >= 2020 else ''

        moeda, valor_original = _detect_moeda(r.get('Valor Sinistrado', ''))
        r['Moeda']          = moeda
        r['Valor Original'] = round(valor_original, 2)

        if d.year >= 2020 and valor_original:
            if moeda == 'BRL':
                pu = _get_ptax(d.year, d.month, 'USD')
                r['Valor BRL'] = round(valor_original, 2)
                r['FX PTAX']   = pu
                r['Valor USD'] = round(valor_original / pu, 2) if pu else 0.0
            elif moeda == 'USD':
                pu = _get_ptax(d.year, d.month, 'USD')
                r['Valor BRL'] = round(valor_original * pu, 2) if pu else 0.0
                r['FX PTAX']   = pu
                r['Valor USD'] = round(valor_original, 2)
            elif moeda == 'EUR':
                pe = _get_ptax(d.year, d.month, 'EUR')
                pu = _get_ptax(d.year, d.month, 'USD')
                vb = round(valor_original * pe, 2) if pe else 0.0
                r['Valor BRL'] = vb
                r['FX PTAX']   = pu
                r['Valor USD'] = round(vb / pu, 2) if pu else 0.0
        else:
            r['Valor BRL'] = 0.0
            r['FX PTAX']   = 0.0
            r['Valor USD'] = 0.0

        r.setdefault('Origem', 'Email')
        r.setdefault('Observações', '')
        r.setdefault('Discricionário', '')

    return registros


# ── Google Sheets ────────────────────────────────────────────

COLUNAS_SHEETS = [
    'ID', 'N° Sinistro', 'Data', 'Data_ISO', 'Mes_Ano',
    'Segurado', 'Filial', 'Apólice',
    'Devedor', 'CNPJ do Devedor', 'Ocorrência', 'Declaração',
    'Valor Sinistrado', 'Moeda', 'Valor Original', 'Valor BRL', 'FX PTAX', 'Valor USD',
    'Vigencia Inicio', 'Vigencia Fim',
    'SETOR', 'SUBSETOR', 'Grupo Econômico',
    'Observações', 'Discricionário', 'Origem',
]
_COLUNAS_NUMERICAS = {'Valor Original', 'Valor BRL', 'FX PTAX', 'Valor USD'}


def escrever_sheets(registros: list):
    if not GSHEET_CREDS or not GSHEET_ID:
        print("Sheets: credenciais não configuradas — pulando.")
        return

    try:
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_info(
            json.loads(GSHEET_CREDS),
            scopes=['https://spreadsheets.google.com/feeds',
                    'https://www.googleapis.com/auth/drive']
        )
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(GSHEET_ID)

        try:
            aba = sh.worksheet('Base')
        except gspread.exceptions.WorksheetNotFound:
            aba = sh.add_worksheet('Base', rows=50000, cols=len(COLUNAS_SHEETS))

        # Ordena por declaração
        todos = sorted(registros, key=_key_decl)

        rows_data = []
        for seq_id, r in enumerate(todos, start=1):
            row = [seq_id]
            for c in COLUNAS_SHEETS[1:]:
                val = r.get(c, '') or ''
                if c == 'Valor Sinistrado':
                    try:
                        row.append(float(r.get('Valor Original', 0.0) or 0.0))
                    except (ValueError, TypeError):
                        row.append(0.0)
                elif c in _COLUNAS_NUMERICAS:
                    try:
                        row.append(float(val))
                    except (ValueError, TypeError):
                        row.append(0.0)
                else:
                    row.append(str(val))
            rows_data.append(row)

        aba.clear()
        titulo = f'Relatório de Sinistros — AVLA Crédito {date.today().year} [BACKFILL {date.today().strftime("%d/%m/%Y")}]'
        aba.update('A1', [[titulo] + [''] * (len(COLUNAS_SHEETS) - 1),
                          COLUNAS_SHEETS] + rows_data,
                   value_input_option='USER_ENTERED')

        print(f"Sheets: {len(todos)} registros gravados (backfill completo).")

        # Formatação básica
        nc  = len(COLUNAS_SHEETS)
        sid = aba.id

        def _rgb(r, g, b):
            return {'red': round(r/255, 4), 'green': round(g/255, 4), 'blue': round(b/255, 4)}

        C_AZUL   = _rgb(0, 48, 135)
        C_BRANCO = _rgb(255, 255, 255)

        def _rng(r0, r1, c0=0, c1=None):
            return {'sheetId': sid, 'startRowIndex': r0, 'endRowIndex': r1,
                    'startColumnIndex': c0, 'endColumnIndex': c1 or nc}

        _fmt = {
            'backgroundColor': C_AZUL,
            'horizontalAlignment': 'CENTER',
            'verticalAlignment': 'MIDDLE',
            'textFormat': {'foregroundColor': C_BRANCO, 'bold': True, 'fontFamily': 'Arial'},
        }
        _fields = 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)'

        requests = [
            {'unmergeCells': {'range': _rng(0, 1)}},
            {'mergeCells': {'range': _rng(0, 1), 'mergeType': 'MERGE_ALL'}},
            {'repeatCell': {'range': _rng(0, 2), 'cell': {'userEnteredFormat': _fmt}, 'fields': _fields}},
            {'updateSheetProperties': {
                'properties': {'sheetId': sid, 'gridProperties': {'frozenRowCount': 2}},
                'fields': 'gridProperties.frozenRowCount',
            }},
        ]
        sh.batch_update({'requests': requests})
        print("Sheets: formatação aplicada.")

    except Exception as e:
        print(f"Erro ao escrever no Sheets: {e}")


# ── Email de confirmação ─────────────────────────────────────

def _logo_path():
    return os.path.join(BASE_DIR, 'avla_logo.png')


def formatar_brl(v: float) -> str:
    if v >= 1_000_000:
        return 'R$ ' + '{:.1f}'.format(v / 1_000_000).replace('.', ',') + 'M'
    if v >= 1_000:
        return 'R$ ' + '{:.0f}'.format(v / 1_000) + 'K'
    return 'R$ ' + '{:.2f}'.format(v).replace('.', ',')


def enviar_email_confirmacao(registros: list, inicio: date, fim: date):
    total    = len(registros)
    total_v  = sum(float(r.get('Valor BRL') or 0) for r in registros)
    logo     = _logo_path()
    tem_logo = os.path.exists(logo)

    img_h = ('<img src="cid:avla_logo" alt="AVLA" style="height:52px;display:block;margin:0 auto;">'
             if tem_logo else '<span style="color:white;font-size:26px;font-weight:bold;">AVLA</span>')

    # Conta por mês
    from collections import Counter
    contagem_mes = Counter(r.get('Mes_Ano', 'sem data') for r in registros)
    linhas_mes = ''.join(
        f'<tr><td style="padding:4px 12px;border-bottom:1px solid #eee;">{m}</td>'
        f'<td style="padding:4px 12px;border-bottom:1px solid #eee;text-align:right;">{n}</td></tr>'
        for m, n in sorted(contagem_mes.items())
    )

    html = (
        '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"></head>'
        '<body style="margin:0;padding:40px 20px;background:#f0f2f5;font-family:Arial,sans-serif;">'
        '<div style="max-width:620px;margin:0 auto;background:#fff;border-radius:4px;overflow:hidden;'
        'box-shadow:0 2px 12px rgba(0,0,0,.12);">'
        f'<div style="background:#0071CE;padding:24px 40px;text-align:center;">{img_h}</div>'
        '<div style="height:6px;background:linear-gradient(to right,#0071CE 0%,#0071CE 30%,'
        '#00A3D9 30%,#00A3D9 55%,#00C4B4 55%,#00C4B4 75%,#7DC242 75%,#7DC242 100%);"></div>'
        '<div style="padding:36px 40px 28px;">'
        '<h1 style="font-size:20px;font-weight:700;color:#0071CE;margin:0 0 6px;">✅ Backfill Concluído</h1>'
        f'<p style="font-size:13px;color:#777;margin:0 0 24px;">'
        f'Período: {inicio.strftime("%d/%m/%Y")} a {fim.strftime("%d/%m/%Y")}</p>'
        '<table style="width:100%;margin-bottom:24px;">'
        '<tr>'
        f'<td style="width:50%;padding:0 8px 0 0;">'
        f'<div style="border:1px solid #e0e0e0;border-top:4px solid #0071CE;border-radius:4px;'
        f'padding:16px;text-align:center;">'
        f'<div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;margin-bottom:8px;">Total de Casos</div>'
        f'<div style="font-size:28px;font-weight:800;color:#0071CE;">{total}</div>'
        f'</div></td>'
        f'<td style="width:50%;padding:0 0 0 8px;">'
        f'<div style="border:1px solid #e0e0e0;border-top:4px solid #00C4B4;border-radius:4px;'
        f'padding:16px;text-align:center;">'
        f'<div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;margin-bottom:8px;">Valor Total (BRL)</div>'
        f'<div style="font-size:22px;font-weight:800;color:#00C4B4;">{formatar_brl(total_v)}</div>'
        f'</div></td>'
        '</tr></table>'
        '<p style="font-size:13px;font-weight:600;color:#333;margin:0 0 8px;">Distribuição por mês:</p>'
        '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:24px;">'
        '<tr style="background:#003087;color:#fff;">'
        '<th style="padding:6px 12px;text-align:left;">Mês</th>'
        '<th style="padding:6px 12px;text-align:right;">Casos</th>'
        '</tr>'
        f'{linhas_mes}'
        '</table>'
        '<div style="background:#f0fff4;border:1px solid #c3e6cb;border-radius:4px;padding:14px 18px;">'
        '<strong style="color:#1D6F42;font-size:13px;">Base do Google Sheets atualizada</strong><br>'
        '<span style="font-size:12px;color:#444;">Todos os registros foram gravados na aba "Base".</span>'
        '</div>'
        '</div></div></body></html>'
    )

    msg = MIMEMultipart('mixed')
    msg['From']    = EMAIL_CAIXA
    msg['To']      = ', '.join(DESTINATARIOS)
    msg['Subject'] = f'[BACKFILL] Sinistros recuperados: {total} casos ({inicio.strftime("%d/%m")} a {fim.strftime("%d/%m/%Y")})'

    related = MIMEMultipart('related')
    related.attach(MIMEText(html, 'html', 'utf-8'))
    if tem_logo:
        img = MIMEImage(open(logo, 'rb').read(), 'png')
        img.add_header('Content-ID', '<avla_logo>')
        img.add_header('Content-Disposition', 'inline', filename='avla_logo.png')
        related.attach(img)

    alt = MIMEMultipart('alternative')
    alt.attach(MIMEText(f"Backfill concluído: {total} sinistros de {inicio} a {fim}.\n", 'plain', 'utf-8'))
    alt.attach(related)
    msg.attach(alt)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as sv:
        sv.starttls()
        sv.login(EMAIL_CAIXA, APP_PASSWORD)
        sv.sendmail(EMAIL_CAIXA, DESTINATARIOS, msg.as_string())

    print(f"Email de confirmação enviado para {', '.join(DESTINATARIOS)}.")


# ── Main ─────────────────────────────────────────────────────

def main():
    hoje = date.today()
    print(f"[{datetime.now().strftime('%d/%m/%Y %H:%M')}] Backfill de sinistros AVLA 2026...")
    print(f"Remetentes monitorados: {', '.join(REMETENTES_VALIDOS)}")

    inicio = datetime(2026, 1, 1, 0, 0, 0)
    fim    = datetime(hoje.year, hoje.month, hoje.day, 23, 59, 59)

    registros = coletar_emails_backfill(inicio, fim)

    if not registros:
        print("Nenhum sinistro encontrado. Verifique os remetentes e credenciais.")
        return

    registros = enriquecer(registros)
    escrever_sheets(registros)
    enviar_email_confirmacao(registros, date(2026, 1, 1), hoje)

    print(f"\nBackfill concluído: {len(registros)} sinistros gravados.")


if __name__ == '__main__':
    main()
