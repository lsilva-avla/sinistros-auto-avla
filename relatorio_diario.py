#!/usr/bin/env python3
"""
Relatório Diário de Sinistros - AVLA Crédito
Coleta emails de aviso de sinistro do dia anterior (D-1), enriquece
com tabelas auxiliares (SEGURADOS_CNAE + GRUPO ECONÔMICO), gera Excel,
escreve na base Google Sheets e envia relatório HTML todo dia às 09h BRT.
"""

import imaplib
import email
import smtplib
import re
import os
import sys
import json
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
EMAIL_CAIXA      = os.environ.get("EMAIL_CAIXA",      "mgignon@avla.com")
APP_PASSWORD     = os.environ.get("APP_PASSWORD",      "")
ASSUNTO_FILTRO   = os.environ.get("ASSUNTO_FILTRO",   "SINISTRO")
REMETENTE_FILTRO = os.environ.get("REMETENTE_FILTRO", "notificaciones-01@avla.com")
DESTINATARIOS    = ["lsilva@avla.com", "mgignon@avla.com"]
GSHEET_CREDS     = os.environ.get("GSHEET_CREDENTIALS", "")
GSHEET_ID        = os.environ.get("GSHEET_ID", "")

IMAP_SERVER = "imap.gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT   = 587

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if not APP_PASSWORD:
    print("ERRO: variável de ambiente APP_PASSWORD não definida.")
    sys.exit(1)
# =============================================================


def periodo_ontem():
    """Retorna (inicio, fim) do dia anterior (D-1)."""
    ontem = date.today() - timedelta(days=1)
    inicio = datetime(ontem.year, ontem.month, ontem.day, 0, 0, 0)
    fim    = datetime(ontem.year, ontem.month, ontem.day, 23, 59, 59)
    return inicio, fim


# ─────────────────────────────────────────────
#  IMAP helpers (iguais ao coletor semanal)
# ─────────────────────────────────────────────

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
                        print(f"✓ Pasta selecionada (All): {nome} — {count} msgs")
                        return
    except Exception as e:
        print(f"Erro ao detectar All Mail: {e}")

    for nome in ['[Gmail]/All Mail', '[Gmail]/Todos os e-mails', 'INBOX']:
        ok, count = _select(mail, nome)
        if ok:
            print(f"✓ Pasta selecionada (fallback): {nome} — {count} msgs")
            return

    raise RuntimeError("Não foi possível selecionar nenhuma pasta de email.")


def _detectar_lixeira(mail):
    try:
        _, pastas = mail.list()
        for item in pastas:
            if item is None:
                continue
            decoded = item.decode('utf-8') if isinstance(item, bytes) else str(item)
            if '\\Trash' in decoded:
                m = re.search(r'"([^"]+)"\s*$|(\S+)\s*$', decoded)
                if m:
                    return (m.group(1) or m.group(2)).strip()
    except Exception as e:
        print(f"Erro ao detectar Lixeira: {e}")
    return None


# ─────────────────────────────────────────────
#  Parsing de email
# ─────────────────────────────────────────────

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


def _buscar_pasta(mail, since, before, vistos):
    try:
        _, ids = mail.search(None, f'(SUBJECT "{ASSUNTO_FILTRO}" SINCE "{since}" BEFORE "{before}")')
    except Exception as e:
        print(f"Erro ao buscar na pasta: {e}")
        return [], vistos

    novos = []
    for msg_id in ids[0].split():
        try:
            _, dados = mail.fetch(msg_id, '(RFC822)')
            msg = email.message_from_bytes(dados[0][1])
        except Exception:
            continue

        mid = msg.get('Message-ID', f'_noid_{msg_id.decode()}_')
        if mid in vistos:
            continue
        vistos.add(mid)

        remetente = str(msg.get('From', ''))
        if REMETENTE_FILTRO and REMETENTE_FILTRO.lower() not in remetente.lower():
            continue

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
            print(f"  ⚠ N° Sinistro não encontrado | assunto: {assunto[:80]}")
        campos['N° Sinistro'] = num_final
        campos['Data']        = data_email.strftime('%d/%m/%Y')
        novos.append(campos)

    return novos, vistos


def coletar_emails(inicio: datetime, fim: datetime) -> list:
    print(f"Conectando em {EMAIL_CAIXA}...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_CAIXA, APP_PASSWORD)

    since  = inicio.strftime("%d-%b-%Y")
    before = (fim + timedelta(days=1)).strftime("%d-%b-%Y")

    vistos    = set()
    registros = []

    selecionar_pasta_completa(mail)
    novos, vistos = _buscar_pasta(mail, since, before, vistos)
    print(f"All Mail: {len(novos)} sinistros")
    registros.extend(novos)

    nome_lixeira = _detectar_lixeira(mail)
    if nome_lixeira:
        ok, count = _select(mail, nome_lixeira)
        if ok:
            print(f"Lixeira ({nome_lixeira}): {count} msgs — verificando...")
            novos, vistos = _buscar_pasta(mail, since, before, vistos)
            print(f"Lixeira: +{len(novos)} sinistros adicionais")
            registros.extend(novos)

    mail.logout()
    print(f"Total coletado: {len(registros)} sinistros")
    return registros


# ─────────────────────────────────────────────
#  Tabelas auxiliares
# ─────────────────────────────────────────────

def carregar_tabelas_auxiliares():
    """Carrega lookups de SEGURADOS_CNAE e GRUPO_ECONOMICO."""
    path_cnae  = os.path.join(BASE_DIR, 'dados', 'SEGURADOS_CNAE.xlsx')
    path_grupo = os.path.join(BASE_DIR, 'dados', 'GRUPO_ECONOMICO.xlsx')

    lookup_cnae  = {}
    lookup_grupo = {}

    # — SEGURADOS_CNAE —
    if os.path.exists(path_cnae):
        try:
            df = pd.read_excel(path_cnae)
            df['poliza'] = df['poliza'].astype(str).str.strip()
            for _, row in df.iterrows():
                vi = row.get('Vigencia Inicio', '')
                vf = row.get('Vigencia Fim', '')
                lookup_cnae[row['poliza']] = {
                    'Vigencia Inicio': vi.strftime('%d/%m/%Y') if hasattr(vi, 'strftime') else str(vi or ''),
                    'Vigencia Fim':    vf.strftime('%d/%m/%Y') if hasattr(vf, 'strftime') else str(vf or ''),
                    'SETOR':           str(row.get('SETOR', '')    or ''),
                    'SUBSETOR':        str(row.get('SUBSETOR', '') or ''),
                }
            print(f"SEGURADOS_CNAE: {len(lookup_cnae)} apólices carregadas.")
        except Exception as e:
            print(f"Erro ao carregar SEGURADOS_CNAE: {e}")
    else:
        print("SEGURADOS_CNAE.xlsx não encontrado — enriquecimento pulado.")

    # — GRUPO ECONÔMICO —
    if os.path.exists(path_grupo):
        try:
            df = pd.read_excel(path_grupo)
            # Coluna 'Identificador Participante' (índice 2) e 'Grupo Econômico' (índice 4)
            col_id    = df.columns[2]   # Identificador Participante
            col_grupo = df.columns[4]   # Grupo Econômico
            for _, row in df.iterrows():
                digits = re.sub(r'\D', '', str(row[col_id] or ''))
                grupo  = str(row[col_grupo] or '').strip()
                if digits and grupo and digits not in lookup_grupo:
                    lookup_grupo[digits] = grupo
            print(f"GRUPO ECONÔMICO: {len(lookup_grupo)} participantes carregados.")
        except Exception as e:
            print(f"Erro ao carregar GRUPO_ECONOMICO: {e}")
    else:
        print("GRUPO_ECONOMICO.xlsx não encontrado — enriquecimento pulado.")

    return lookup_cnae, lookup_grupo


def enriquecer(registros: list, lookup_cnae: dict, lookup_grupo: dict) -> list:
    """Adiciona colunas de CNAE e Grupo Econômico a cada registro."""
    for r in registros:
        # — SEGURADOS_CNAE pelo número da apólice —
        apolice   = str(r.get('Apólice', '')).strip()
        cnae_info = lookup_cnae.get(apolice, {})
        r['Vigencia Inicio'] = cnae_info.get('Vigencia Inicio', '')
        r['Vigencia Fim']    = cnae_info.get('Vigencia Fim', '')
        r['SETOR']           = cnae_info.get('SETOR', '')
        r['SUBSETOR']        = cnae_info.get('SUBSETOR', '')

        # — GRUPO ECONÔMICO pelo CNPJ (só dígitos) —
        cnpj_digits         = re.sub(r'\D', '', str(r.get('CNPJ do Devedor', '')))
        r['Grupo Econômico'] = lookup_grupo.get(cnpj_digits, '')

    return registros


# ─────────────────────────────────────────────
#  Google Sheets
# ─────────────────────────────────────────────

COLUNAS_SHEETS = [
    'ID', 'N° Sinistro', 'Data', 'Segurado', 'Filial', 'Apólice',
    'Devedor', 'CNPJ do Devedor', 'Ocorrência', 'Declaração',
    'Valor Sinistrado', 'Vigencia Inicio', 'Vigencia Fim',
    'SETOR', 'SUBSETOR', 'Grupo Econômico',
]


def escrever_sheets(registros: list):
    if not GSHEET_CREDS or not GSHEET_ID:
        print("Sheets: GSHEET_CREDENTIALS ou GSHEET_ID não configurados — pulando.")
        return

    try:
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_info(
            json.loads(GSHEET_CREDS),
            scopes=[
                'https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive',
            ]
        )
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(GSHEET_ID)

        try:
            aba = sh.worksheet('Base')
        except gspread.exceptions.WorksheetNotFound:
            aba = sh.add_worksheet('Base', rows=50000, cols=len(COLUNAS_SHEETS))

        valores = aba.get_all_values()

        if not valores:
            aba.append_row(COLUNAS_SHEETS, value_input_option='USER_ENTERED')
            existentes = set()
            prox_id    = 1
        else:
            header    = valores[0]
            idx_num   = header.index('N° Sinistro') if 'N° Sinistro' in header else 1
            existentes = {row[idx_num] for row in valores[1:] if len(row) > idx_num}
            prox_id   = len(valores)   # ID sequencial continua de onde parou

        novos = [r for r in registros if str(r.get('N° Sinistro', '')) not in existentes]

        if not novos:
            print(f"Sheets: nenhum registro novo (já existem {len(existentes)}).")
            return

        rows = []
        for i, r in enumerate(novos, start=prox_id):
            row = [i] + [str(r.get(c, '') or '') for c in COLUNAS_SHEETS[1:]]
            rows.append(row)

        aba.append_rows(rows, value_input_option='USER_ENTERED')
        print(f"Sheets: {len(novos)} novos registros inseridos (total agora: {prox_id + len(novos) - 1}).")

    except Exception as e:
        print(f"Erro ao escrever no Sheets (não fatal): {e}")


# ─────────────────────────────────────────────
#  Excel
# ─────────────────────────────────────────────

def gerar_excel(registros: list, ontem: date) -> BytesIO:
    colunas = [
        'ID', 'N° Sinistro', 'Data', 'Segurado', 'Filial', 'Apólice',
        'Devedor', 'CNPJ do Devedor', 'Ocorrência', 'Declaração',
        'Valor Sinistrado', 'Vigencia Inicio', 'Vigencia Fim',
        'SETOR', 'SUBSETOR', 'Grupo Econômico',
    ]

    df = pd.DataFrame(registros)
    df.insert(0, 'ID', range(1, len(df) + 1))
    for col in colunas:
        if col not in df.columns:
            df[col] = ''
    df = df[colunas]

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        nome_aba = f"Sinistros {ontem.strftime('%d-%m-%Y')}"
        df.to_excel(writer, index=False, sheet_name=nome_aba)
        ws = writer.sheets[nome_aba]

        # Cabeçalho
        header_font = Font(bold=True, color='FFFFFF', size=11)
        header_fill = PatternFill(fill_type='solid', fgColor='003087')
        for cell in ws[1]:
            cell.font      = header_font
            cell.fill      = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 20

        # Colunas enriquecidas com cabeçalho destacado em verde escuro
        cols_enrich = ['Vigencia Inicio', 'Vigencia Fim', 'SETOR', 'SUBSETOR', 'Grupo Econômico']
        enrich_fill = PatternFill(fill_type='solid', fgColor='1D6F42')
        for cell in ws[1]:
            if cell.value in cols_enrich:
                cell.fill = enrich_fill

        # Auto-width
        for col in ws.columns:
            max_len = max((len(str(c.value or '')) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 55)

        # Destaque: ontem em azul claro (= todos os registros, já que são de ontem)
        destaque = PatternFill(fill_type='solid', fgColor='DCF0FF')
        ontem_str = ontem.strftime('%d/%m/%Y')
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            try:
                if str(row[2].value) == ontem_str:
                    for cell in row:
                        cell.fill = destaque
            except (ValueError, TypeError):
                pass

    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
#  Email HTML
# ─────────────────────────────────────────────

def _logo_path():
    return os.path.join(BASE_DIR, 'avla_logo.png')


def parse_valor(s: str) -> float:
    try:
        s = re.sub(r'[^\d,]', '', str(s))
        return float(s.replace(',', '.')) if s else 0.0
    except Exception:
        return 0.0


def formatar_brl(v: float) -> str:
    if v >= 1_000_000:
        return 'R$ ' + '{:.1f}'.format(v / 1_000_000).replace('.', ',') + 'M'
    if v >= 1_000:
        return 'R$ ' + '{:.0f}'.format(v / 1_000) + 'K'
    return 'R$ ' + '{:.2f}'.format(v).replace('.', ',')


def gerar_html_email(qtd, registros, label_periodo, nome_arquivo, tem_logo=True):
    img_h = ('<img src="cid:avla_logo" alt="AVLA" style="height:52px;display:block;margin:0 auto;">'
             if tem_logo else '<span style="color:white;font-size:26px;font-weight:bold;">AVLA</span>')
    img_f = ('<img src="cid:avla_logo" alt="AVLA" style="height:36px;display:block;margin:0 auto;">'
             if tem_logo else '<span style="color:white;font-size:18px;font-weight:bold;">AVLA</span>')

    total_v = sum(parse_valor(r.get('Valor Sinistrado', '')) for r in registros)
    maior_v = max((parse_valor(r.get('Valor Sinistrado', '')) for r in registros), default=0.0)
    maior_r = next((r for r in registros if parse_valor(r.get('Valor Sinistrado', '')) == maior_v), {})
    maior_nome = maior_r.get('Devedor', maior_r.get('Segurado', '—'))
    if len(maior_nome) > 22:
        maior_nome = maior_nome[:22] + '…'

    # Enriquecimento: quantos tiveram grupo econômico e setor identificados
    com_grupo = sum(1 for r in registros if r.get('Grupo Econômico', ''))
    com_setor = sum(1 for r in registros if r.get('SETOR', ''))

    def card(cor, label, val, sub, width='33%'):
        fs = '22' if label == 'Total de Casos' else '18'
        return (
            f'<td width="{width}" style="padding:0 5px;">'
            f'<div style="border:1px solid #e0e0e0;border-top:4px solid {cor};border-radius:4px;'
            f'padding:16px 12px 12px;text-align:center;">'
            f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;'
            f'letter-spacing:.5px;color:#888;margin-bottom:8px;">{label}</div>'
            f'<div style="font-size:{fs}px;font-weight:800;color:#1a1a1a;line-height:1.1;">{val}</div>'
            f'<div style="font-size:10px;color:#aaa;margin-top:4px;">{sub}</div>'
            f'</div></td>'
        )

    return (
        '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"></head>'
        '<body style="margin:0;padding:40px 20px;background:#f0f2f5;font-family:Arial,Helvetica,sans-serif;">'
        '<div style="max-width:620px;margin:0 auto;background:#fff;border-radius:4px;overflow:hidden;'
        'box-shadow:0 2px 12px rgba(0,0,0,.12);">'
        f'<div style="background:#0071CE;padding:24px 40px;text-align:center;">{img_h}</div>'
        '<div style="height:6px;background:linear-gradient(to right,#0071CE 0%,#0071CE 30%,'
        '#00A3D9 30%,#00A3D9 55%,#00C4B4 55%,#00C4B4 75%,#7DC242 75%,#7DC242 100%);"></div>'
        '<div style="padding:36px 40px 28px;">'
        '<p style="font-size:14px;color:#444;margin:0 0 8px;">Olá, equipe de Crédito,</p>'
        '<h1 style="font-size:22px;font-weight:700;color:#0071CE;margin:0 0 6px;">'
        'Relatório Diário de Sinistros</h1>'
        f'<p style="font-size:13px;color:#777;margin:0 0 24px;">Período: {label_periodo}</p>'
        '<p style="font-size:14px;color:#444;line-height:1.6;margin:0 0 28px;">'
        'Segue o resumo dos avisos de sinistro registrados ontem. '
        'Os dados completos — incluindo vigência da apólice, setor e grupo econômico — '
        'estão na planilha Excel em anexo.</p>'
        '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:12px;"><tr>'
        f'{card("#0071CE","Total de Casos",qtd,"sinistros ontem")}'
        f'{card("#00C4B4","Valor Total",formatar_brl(total_v),"soma dos sinistrados")}'
        f'{card("#7DC242","Maior Caso",formatar_brl(maior_v),maior_nome)}'
        '</tr></table>'
        '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:28px;"><tr>'
        f'{card("#00A3D9","Com Grupo Econ.",com_grupo,"devedores identificados","50%")}'
        f'{card("#1D6F42","Com Setor CNAE",com_setor,"apólices identificadas","50%")}'
        '</tr></table>'
        '<div style="background:#f5f8ff;border:1px solid #d0dff5;border-radius:4px;'
        'padding:14px 18px;margin-bottom:28px;">'
        f'<strong style="color:#1D6F42;display:block;font-size:13px;margin-bottom:4px;">{nome_arquivo}</strong>'
        f'<span style="font-size:13px;color:#444;">Planilha com todos os {qtd} casos · '
        'Nº Sinistro, Data, Segurado, Devedor, CNPJ, Valor, Setor, Grupo Econômico</span>'
        '</div>'
        '<hr style="border:none;border-top:1px solid #eee;margin-bottom:20px;">'
        '<p style="font-size:13px;color:#666;line-height:1.7;margin:0;">'
        'Este relatório é gerado automaticamente todo dia às 09h (BRT).<br>'
        'Os dados também são registrados na base online Google Sheets.</p>'
        '</div>'
        f'<div style="background:#0071CE;padding:20px 40px;text-align:center;">{img_f}'
        '<div style="font-size:11px;color:rgba(255,255,255,.6);margin-top:10px;">'
        'Mensagem automática — não é necessário responder</div>'
        '</div>'
        '</div></body></html>'
    )


def enviar_relatorio(buf: BytesIO, qtd: int, registros: list, ontem: date):
    label_periodo = ontem.strftime('%d/%m/%Y')
    nome_arquivo  = f"Sinistros_Diario_{ontem.strftime('%d%m%Y')}.xlsx"

    logo     = _logo_path()
    tem_logo = os.path.exists(logo)

    plain_text = (
        f"Olá equipe,\n\nSegue o relatório diário de sinistros de {label_periodo}.\n"
        f"Total: {qtd} sinistros.\n\nGerado automaticamente às 09h BRT.\n"
    )

    msg = MIMEMultipart('mixed')
    msg['From']    = EMAIL_CAIXA
    msg['To']      = ', '.join(DESTINATARIOS)
    msg['Subject'] = f'Relatório Diário de Sinistros — {label_periodo} ({qtd} casos)'

    related = MIMEMultipart('related')
    related.attach(MIMEText(
        gerar_html_email(qtd, registros, label_periodo, nome_arquivo, tem_logo),
        'html', 'utf-8'
    ))
    if tem_logo:
        img = MIMEImage(open(logo, 'rb').read(), 'png')
        img.add_header('Content-ID', '<avla_logo>')
        img.add_header('Content-Disposition', 'inline', filename='avla_logo.png')
        related.attach(img)

    alt = MIMEMultipart('alternative')
    alt.attach(MIMEText(plain_text, 'plain', 'utf-8'))
    alt.attach(related)
    msg.attach(alt)

    parte = MIMEBase('application', 'octet-stream')
    parte.set_payload(buf.read())
    encoders.encode_base64(parte)
    parte.add_header('Content-Disposition', f'attachment; filename="{nome_arquivo}"')
    msg.attach(parte)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as servidor:
        servidor.starttls()
        servidor.login(EMAIL_CAIXA, APP_PASSWORD)
        servidor.sendmail(EMAIL_CAIXA, DESTINATARIOS, msg.as_string())

    print(f"Email enviado para {', '.join(DESTINATARIOS)} — {qtd} sinistros.")


# ─────────────────────────────────────────────
#  Main
# ─────────────────────────────────────────────

def main():
    agora = datetime.now()
    print(f"[{agora.strftime('%d/%m/%Y %H:%M')}] Iniciando relatório diário AVLA...")

    inicio, fim = periodo_ontem()
    ontem = inicio.date()
    print(f"Período: {ontem.strftime('%d/%m/%Y')} (D-1)")

    registros = coletar_emails(inicio, fim)

    if not registros:
        print("Nenhum sinistro ontem. Email não enviado.")
        return

    # Enriquece com tabelas auxiliares
    lookup_cnae, lookup_grupo = carregar_tabelas_auxiliares()
    registros = enriquecer(registros, lookup_cnae, lookup_grupo)

    # Escreve na base online
    escrever_sheets(registros)

    # Gera Excel e envia email
    excel = gerar_excel(registros, ontem)
    enviar_relatorio(excel, len(registros), registros, ontem)

    print("Concluído com sucesso.")


if __name__ == '__main__':
    main()
