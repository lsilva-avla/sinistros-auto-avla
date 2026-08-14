#!/usr/bin/env python3
"""
Diagnóstico IMAP — AVLA Crédito
Conecta no Gmail e lista os emails de notificaciones-01@avla.com
dos últimos 90 dias, SEM filtro de assunto, para identificar o que está chegando.
"""

import imaplib
import email
import re
import os
import sys
from datetime import datetime, timedelta, date

EMAIL_CAIXA      = os.environ.get("EMAIL_CAIXA",      "mgignon@avla.com")
APP_PASSWORD     = os.environ.get("APP_PASSWORD",      "")
REMETENTE_FILTRO = os.environ.get("REMETENTE_FILTRO", "notificaciones-01@avla.com")

IMAP_SERVER = "imap.gmail.com"

if not APP_PASSWORD:
    print("ERRO: APP_PASSWORD não definido.")
    sys.exit(1)


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


def main():
    print(f"[{datetime.now().strftime('%d/%m/%Y %H:%M')}] Diagnóstico IMAP AVLA...")
    print(f"Conta: {EMAIL_CAIXA}")
    print(f"Remetente filtrado: {REMETENTE_FILTRO}")
    print()

    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_CAIXA, APP_PASSWORD)
    print("✓ Login OK")

    # Lista todas as pastas
    _, pastas = mail.list()
    print("\n=== Pastas disponíveis ===")
    for item in pastas:
        if item:
            print(" ", item.decode('utf-8') if isinstance(item, bytes) else item)
    print("==========================\n")

    # Seleciona All Mail
    pasta_selecionada = None
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
                    print(f"✓ All Mail selecionado: {nome} — {count} msgs total")
                    pasta_selecionada = nome
                    break

    if not pasta_selecionada:
        for nome in ['[Gmail]/All Mail', '[Gmail]/Todos os e-mails', 'INBOX']:
            ok, count = _select(mail, nome)
            if ok:
                print(f"✓ Pasta fallback: {nome} — {count} msgs")
                pasta_selecionada = nome
                break

    # Busca 1: SEM filtro de assunto, últimos 90 dias, do remetente esperado
    since_90 = (date.today() - timedelta(days=90)).strftime("%d-%b-%Y")
    before_amanha = (date.today() + timedelta(days=1)).strftime("%d-%b-%Y")

    print(f"\n--- Busca 1: FROM '{REMETENTE_FILTRO}' (90 dias, sem filtro assunto) ---")
    _, ids = mail.search(None, f'(FROM "{REMETENTE_FILTRO}" SINCE "{since_90}" BEFORE "{before_amanha}")')
    ids_list = ids[0].split() if ids and ids[0] else []
    print(f"Emails encontrados: {len(ids_list)}")

    for msg_id in ids_list[-20:]:   # últimos 20
        try:
            _, dados = mail.fetch(msg_id, '(BODY[HEADER.FIELDS (SUBJECT DATE FROM)])')
            raw = dados[0][1].decode('utf-8', errors='ignore') if dados and dados[0] else ''
            assunto = ''
            data_hdr = ''
            remetente = ''
            for linha in raw.splitlines():
                if linha.lower().startswith('subject:'):
                    assunto = linha[8:].strip()
                elif linha.lower().startswith('date:'):
                    data_hdr = linha[5:].strip()
                elif linha.lower().startswith('from:'):
                    remetente = linha[5:].strip()
            print(f"  [{data_hdr[:16]}] FROM: {remetente[:50]} | ASSUNTO: {assunto[:80]}")
        except Exception as e:
            print(f"  Erro ao ler cabeçalho: {e}")

    # Busca 2: COM filtro de assunto SINISTRO
    print(f"\n--- Busca 2: SUBJECT 'SINISTRO' (90 dias) ---")
    _, ids2 = mail.search(None, f'(SUBJECT "SINISTRO" SINCE "{since_90}" BEFORE "{before_amanha}")')
    ids2_list = ids2[0].split() if ids2 and ids2[0] else []
    print(f"Com SUBJECT 'SINISTRO': {len(ids2_list)} emails")

    # Busca 3: COM filtro SINIESTRO (espanhol com 'e')
    print(f"\n--- Busca 3: SUBJECT 'SINIESTRO' (90 dias) ---")
    _, ids3 = mail.search(None, f'(SUBJECT "SINIESTRO" SINCE "{since_90}" BEFORE "{before_amanha}")')
    ids3_list = ids3[0].split() if ids3 and ids3[0] else []
    print(f"Com SUBJECT 'SINIESTRO': {len(ids3_list)} emails")

    # Busca 4: AVISO (termo genérico)
    print(f"\n--- Busca 4: SUBJECT 'AVISO' (90 dias, do remetente) ---")
    _, ids4 = mail.search(None, f'(FROM "{REMETENTE_FILTRO}" SUBJECT "AVISO" SINCE "{since_90}" BEFORE "{before_amanha}")')
    ids4_list = ids4[0].split() if ids4 and ids4[0] else []
    print(f"Com SUBJECT 'AVISO': {len(ids4_list)} emails")

    # Resumo
    print("\n=== RESUMO ===")
    print(f"Emails do remetente (90d):    {len(ids_list)}")
    print(f"Com SINISTRO no assunto:      {len(ids2_list)}")
    print(f"Com SINIESTRO no assunto:     {len(ids3_list)}")
    print(f"Com AVISO no assunto:         {len(ids4_list)}")

    if len(ids_list) > 0 and len(ids2_list) == 0 and len(ids3_list) == 0:
        print("\n⚠️  DIAGNÓSTICO: Emails chegando do remetente, mas assunto mudou!")
        print("   → Verifique o assunto dos emails acima e ajuste o ASSUNTO_FILTRO no GitHub Secret.")
    elif len(ids_list) == 0:
        print("\n⚠️  DIAGNÓSTICO: Nenhum email do remetente nos últimos 90 dias!")
        print("   → O endereço remetente pode ter mudado, ou emails não estão chegando.")
    else:
        print("\n✓  Emails encontrados com filtros atuais — problema pode ser de parse.")

    mail.logout()
    print("\nDiagnóstico concluído.")


if __name__ == '__main__':
    main()
