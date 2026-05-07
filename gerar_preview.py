import base64

with open('logo_b64.txt', 'r') as f:
    b64 = f.read().strip()

html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Preview - Email Relatorio Semanal de Sinistros</title>
<style>
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ background: #f0f2f5; font-family: Arial, Helvetica, sans-serif; padding: 40px 20px; }}
  .email-wrapper {{ max-width: 620px; margin: 0 auto; background: #ffffff; border-radius: 4px; overflow: hidden; box-shadow: 0 2px 12px rgba(0,0,0,0.12); }}
  .header {{ background: #0071CE; padding: 24px 40px; text-align: center; }}
  .stripe {{ height: 6px; background: linear-gradient(to right, #0071CE 0%, #0071CE 30%, #00A3D9 30%, #00A3D9 55%, #00C4B4 55%, #00C4B4 75%, #7DC242 75%, #7DC242 100%); }}
  .body {{ padding: 36px 40px 28px; }}
  .greeting {{ font-size: 14px; color: #444; margin-bottom: 8px; }}
  .title {{ font-size: 22px; font-weight: 700; color: #0071CE; margin-bottom: 6px; }}
  .period {{ font-size: 13px; color: #777; margin-bottom: 24px; }}
  .intro {{ font-size: 14px; color: #444; line-height: 1.6; margin-bottom: 28px; }}
  .cards {{ display: flex; gap: 14px; margin-bottom: 28px; }}
  .card {{ flex: 1; border: 1px solid #e0e0e0; border-top: 4px solid #0071CE; border-radius: 4px; padding: 18px 16px 14px; text-align: center; }}
  .card.teal {{ border-top-color: #00C4B4; }}
  .card.green {{ border-top-color: #7DC242; }}
  .card-label {{ font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: #888; margin-bottom: 10px; }}
  .card-value {{ font-size: 22px; font-weight: 800; color: #1a1a1a; line-height: 1.1; }}
  .card-sub {{ font-size: 11px; color: #aaa; margin-top: 4px; }}
  .attachment {{ background: #f5f8ff; border: 1px solid #d0dff5; border-radius: 4px; padding: 14px 18px; display: flex; align-items: center; gap: 12px; margin-bottom: 28px; }}
  .attachment-icon {{ width: 36px; height: 36px; background: #1D6F42; border-radius: 4px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }}
  .attachment-icon svg {{ width: 20px; height: 20px; fill: white; }}
  .attachment-text {{ font-size: 13px; color: #444; }}
  .attachment-text strong {{ color: #1D6F42; display: block; font-size: 13px; }}
  hr {{ border: none; border-top: 1px solid #eee; margin-bottom: 20px; }}
  .closing {{ font-size: 13px; color: #666; line-height: 1.7; }}
  .footer {{ background: #0071CE; padding: 20px 40px; text-align: center; }}
  .footer-auto {{ font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 10px; }}
</style>
</head>
<body>
<div class="email-wrapper">

  <div class="header">
    <img src="data:image/png;base64,{b64}" alt="Avla" style="height:52px; display:block; margin:0 auto;">
  </div>
  <div class="stripe"></div>

  <div class="body">
    <p class="greeting">Ol&aacute;, equipe de Cr&eacute;dito,</p>
    <h1 class="title">Relat&oacute;rio Semanal de Sinistros</h1>
    <p class="period">Per&iacute;odo: 27/04/2026 (dom) a 03/05/2026 (s&aacute;b)</p>
    <p class="intro">
      Segue o resumo dos avisos de sinistro registrados na semana.
      Os dados completos est&atilde;o na planilha Excel em anexo, com todas as
      informa&ccedil;&otilde;es por caso: segurado, devedor, ap&oacute;lice, CNPJ e valores.
    </p>
    <div class="cards">
      <div class="card">
        <div class="card-label">Total de Casos</div>
        <div class="card-value">12</div>
        <div class="card-sub">sinistros na semana</div>
      </div>
      <div class="card teal">
        <div class="card-label">Valor Total</div>
        <div class="card-value" style="font-size:18px;">R$ 8,4M</div>
        <div class="card-sub">soma dos sinistrados</div>
      </div>
      <div class="card green">
        <div class="card-label">Maior Caso</div>
        <div class="card-value" style="font-size:18px;">R$ 1,2M</div>
        <div class="card-sub">Qu&iacute;mica Amparo Ltda</div>
      </div>
    </div>
    <div class="attachment">
      <div class="attachment-icon">
        <svg viewBox="0 0 24 24"><path d="M19 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V5a2 2 0 0 0-2-2zm-7 3l5 5h-3v4h-4v-4H7l5-5z"/></svg>
      </div>
      <div class="attachment-text">
        <strong>Sinistros_2704_03052026.xlsx</strong>
        Planilha com todos os 12 casos &middot; colunas: ID, N&deg; Sinistro, Data, Segurado, Filial, Ap&oacute;lice, Devedor, CNPJ, Ocorr&ecirc;ncia, Declara&ccedil;&atilde;o, Valor
      </div>
    </div>
    <hr>
    <p class="closing">
      Este relat&oacute;rio &eacute; gerado automaticamente toda sexta-feira.<br>
      Em caso de d&uacute;vidas, responda a este email.
    </p>
  </div>

  <div class="footer">
    <img src="data:image/png;base64,{b64}" alt="Avla" style="height:36px; display:block; margin:0 auto;">
    <div class="footer-auto">Mensagem autom&aacute;tica &mdash; n&atilde;o &eacute; necess&aacute;rio responder</div>
  </div>

</div>
</body>
</html>"""

with open('preview_email.html', 'w', encoding='utf-8') as f:
    f.write(html)
print('preview_email.html gerado com sucesso')
