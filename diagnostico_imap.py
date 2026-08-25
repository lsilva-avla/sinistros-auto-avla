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

    since_90   = (date.today() - timedelta(days=90)).strftime("%d-%b-%Y")
    since_ago1 = date(2026, 8, 1).strftime("%d-%b-%Y")   # 1/ago em diante
    before_amanha = (date.today() + timedelta(days=1)).strftime("%d-%b-%Y")

    def listar_emails(ids_list, label, limite=30):
        print(f"  ({len(ids_list)} encontrados)")
        for msg_id in ids_list[-limite:]:
            try:
                _, dados = mail.fetch(msg_id, '(BODY[HEADER.FIELDS (SUBJECT DATE FROM)])')
                raw = dados[0][1].decode('utf-8', errors='ignore') if dados and dados[0] else ''
                assunto = remetente = data_hdr = ''
                for linha in raw.splitlines():
                    l = linha.lower()
                    if l.startswith('subject:'):  assunto   = linha[8:].strip()
                    elif l.startswith('date:'):   data_hdr  = linha[5:].strip()
                    elif l.startswith('from:'):   remetente = linha[5:].strip()
                print(f"    [{data_hdr[:16]}] FROM: {remetente[:55]} | {assunto[:70]}")
            except Exception as e:
                print(f"    Erro: {e}")

    # ── BLOCO 1: remetente esperado — 90 dias ───────────────────
    print(f"\n--- [1] FROM '{REMETENTE_FILTRO}' — últimos 90 dias ---")
    _, r1 = mail.search(None, f'(FROM "{REMETENTE_FILTRO}" SINCE "{since_90}" BEFORE "{before_amanha}")')
    l1 = r1[0].split() if r1 and r1[0] else []
    listar_emails(l1, "remetente 90d")

    # ── BLOCO 2: remetente esperado — desde 1/ago ───────────────
    print(f"\n--- [2] FROM '{REMETENTE_FILTRO}' — desde 01/ago/2026 ---")
    _, r2 = mail.search(None, f'(FROM "{REMETENTE_FILTRO}" SINCE "{since_ago1}" BEFORE "{before_amanha}")')
    l2 = r2[0].split() if r2 and r2[0] else []
    listar_emails(l2, "remetente ago")

    # ── BLOCO 3: SUBJECT SINISTRO — desde 1/ago, qualquer remetente ─
    print(f"\n--- [3] SUBJECT 'SINISTRO' — desde 01/ago, qualquer remetente ---")
    _, r3 = mail.search(None, f'(SUBJECT "SINISTRO" SINCE "{since_ago1}" BEFORE "{before_amanha}")')
    l3 = r3[0].split() if r3 and r3[0] else []
    listar_emails(l3, "sinistro ago")

    # ── BLOCO 4: SUBJECT SINISTRO — sem limite de data ──────────
    print(f"\n--- [4] SUBJECT 'SINISTRO' — sem limite de data (últimos 10) ---")
    _, r4 = mail.search(None, 'SUBJECT "SINISTRO"')
    l4 = r4[0].split() if r4 and r4[0] else []
    listar_emails(l4, "sinistro all", limite=10)

    # ── RESUMO ───────────────────────────────────────────────────
    print("\n=== RESUMO ===")
    print(f"FROM remetente (90d):         {len(l1)}")
    print(f"FROM remetente (desde 01/ago):{len(l2)}")
    print(f"SINISTRO no assunto (01/ago+):{len(l3)}")
    print(f"SINISTRO no assunto (total):  {len(l4)}")

    if len(l2) == 0 and len(l3) == 0:
        print("\n⚠️  DIAGNÓSTICO CRÍTICO: nenhum email de sinistro chegou desde 01/ago!")
        print("   → O sistema parou de enviar notificações para esta caixa.")
        print("   → Verificar se o endereço receptor mudou no portal de sinistros.")
    elif len(l2) == 0 and len(l3) > 0:
        print("\n⚠️  DIAGNÓSTICO: Emails com SINISTRO chegando de OUTRO remetente desde 01/ago!")
        print("   → O sender mudou. Ajustar REMETENTE_FILTRO no GitHub Secret.")
    elif len(l2) > 0:
        print(f"\n✓  Emails do remetente chegando normalmente ({len(l2)} desde 01/ago).")
        print("   → O problema é de timing (relatório semanal com lag) ou de parse.")

    mail.logout()
    print("\nDiagnóstico concluído.")


if __name__ == '__main__':
    main()
