# Configuração — Coletor Semanal de Sinistros AVLA

## Como funciona

O script lê emails da conta **mgignon@avla.com** (onde chegam os avisos de sinistro),
extrai os dados e envia o relatório Excel toda **sexta-feira às 08h (BRT)** de volta para
**mgignon@avla.com**. Roda no GitHub Actions — nenhum
computador precisa estar ligado.

---

## 1. Instalar dependências (uma única vez, para testes locais)

```
pip install -r requirements.txt
```

---

## 2. App Password — já obtida ✅

A senha de 16 caracteres foi gerada em **mgignon@avla.com**.
Guarde-a em local seguro — ela será usada no passo 3.

---

## 3. Setup no GitHub (5 passos)

### Criar repositório privado
1. github.com → **New repository**
2. Nome: `sinistros-auto-avla`
3. Marque **Private**
4. Clique **Create repository**

### Subir os arquivos
No PowerShell, dentro da pasta `sinistros-auto/`:
```powershell
git init
git add .
git commit -m "feat: coletor semanal e mensal de sinistros"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/sinistros-auto-avla.git
git push -u origin main
```

### Adicionar os 4 Secrets
**Settings → Secrets and variables → Actions → New repository secret**

| Secret           | Valor                          |
|------------------|--------------------------------|
| `EMAIL_CAIXA`    | `mgignon@avla.com`             |
| `APP_PASSWORD`   | senha de 16 chars gerada       |
| `REMETENTE_FILTRO` | `sub.credito@avla.com`       |
| `EMAIL_DESTINO`  | `mgignon@avla.com`             |

### Testar manualmente
Aba **Actions** → workflow → **"Run workflow"** → aguardar ~30s → verificar log.
Se verde: o email chegou em sub.credito@avla.com com o Excel em anexo. ✅

---

## Agendamentos ativos

| Script                  | Quando                        |
|-------------------------|-------------------------------|
| `coletor_sinistros.py`  | Toda sexta, 08h BRT           |
| `relatorio_mensal.py`   | Todo dia 1º do mês, 08h BRT  |

---

## Estrutura de arquivos

```
sinistros-auto/
├── coletor_sinistros.py              ← relatório semanal
├── relatorio_mensal.py               ← relatório mensal
├── requirements.txt
├── executar.bat                      ← rodar localmente
├── avla_logo.png
├── preview_email.html                ← mock visual do email
├── CONFIGURACAO.md                   ← este guia
└── .github/workflows/
    ├── relatorio_semanal.yml         ← GitHub Actions sexta 08h
    └── relatorio_mensal.yml          ← GitHub Actions dia 1º 08h
```
