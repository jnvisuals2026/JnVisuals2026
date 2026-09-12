JN VISUALS — VERSÃO VERCEL

1. No Vercel, importa esta pasta/repositório.
2. O index.html e logo.png estão na raiz.
3. O formulário usa /api/submit.
4. No Vercel, cria uma Environment Variable:
   Nome: APPS_SCRIPT_URL
   Valor: o URL /exec do Google Apps Script.
5. Faz redeploy.

GOOGLE APPS SCRIPT
- Abre script.google.com.
- Cria um projeto e coloca o conteúdo de apps-script/Code.gs.
- Deploy > New deployment > Web app.
- Execute as: Me.
- Who has access: Anyone.
- Copia o URL terminado em /exec.
- Coloca esse URL na variável APPS_SCRIPT_URL do Vercel.

EMAIL
As notificações de aprovação são enviadas para:
jnvisuals2026@gmail.com

O Google Apps Script usa uma Google Sheet chamada "Adesoes" para guardar os pedidos.

NOTA
O Vercel aloja o site e a API. O envio de Gmail é feito pelo Google Apps Script. Não coloques uma palavra-passe do Gmail no código do site.
