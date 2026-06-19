# G5 Instrumentos Controle

Aplicação independente para controle de instrumentos, equipes, equipamentos, materiais, compras, obras e pendências usando Issues do Autodesk Construction Cloud / Autodesk Platform Services.

Este projeto foi separado da Central G5 para uso próprio do setor de Instrumentos.

## Deploy no Render

Configuração recomendada no Render:

- **Language:** Node
- **Branch:** main
- **Build Command:** `npm install && npm run build`
- **Start Command:** `npm start`

## Variáveis de ambiente

Configure no Render, sem salvar valores reais no GitHub:

```env
APS_CLIENT_ID=client_id_do_app_instrumentos
APS_CLIENT_SECRET=client_secret_do_app_instrumentos
APS_CALLBACK_URL=https://g5-instrumentos-controle.onrender.com/api/auth/callback
APS_SCOPES=data:read data:write data:create account:read
NODE_ENV=production
SESSION_SECRET=uma_frase_grande_e_aleatoria
SESSION_MAX_AGE_MS=1209600000
```

No app APS **Instrumentos**, cadastre o mesmo callback:

```txt
https://g5-instrumentos-controle.onrender.com/api/auth/callback
```

## Desenvolvimento local

```bash
npm install
npm run dev
```

Para rodar apenas o servidor em produção local:

```bash
npm run build
npm start
```

## Observações de segurança

- Nunca salve `APS_CLIENT_SECRET` no código.
- Use as variáveis de ambiente do Render.
- Se o segredo já foi exposto em print ou conversa, regenere no painel Autodesk APS.
