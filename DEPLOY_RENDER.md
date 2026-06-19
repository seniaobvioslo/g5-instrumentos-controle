# Deploy Render - G5 Instrumentos Controle

1. Crie/abra o serviço `g5-instrumentos-controle` no Render.
2. Use este repositório na branch `main`.
3. Configure:
   - Language: Node
   - Build Command: `npm install && npm run build`
   - Start Command: `npm start`
4. Configure as variáveis de ambiente:
   - `APS_CLIENT_ID`
   - `APS_CLIENT_SECRET`
   - `APS_CALLBACK_URL=https://g5-instrumentos-controle.onrender.com/api/auth/callback`
   - `APS_SCOPES=data:read data:write data:create account:read`
   - `NODE_ENV=production`
   - `SESSION_SECRET`
5. No APS, cadastre o mesmo Callback URL.
6. Faça `Manual Deploy > Deploy latest commit`.

Teste de saúde:

```txt
https://g5-instrumentos-controle.onrender.com/health
```
