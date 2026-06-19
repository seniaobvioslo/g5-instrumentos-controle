# Publicar e conectar no Autodesk Construction Cloud / Data Management

Este guia explica o caminho para tirar o app do `localhost` e disponibilizar o **G5 - Gestão e coordenação** como uma integração acessível pela empresa.

## O ponto mais importante

O Autodesk Data Management / ACC não consegue abrir um app que está em:

```txt
http://localhost:5173
```

Esse endereço só existe no seu computador.

Para outras pessoas e para o ambiente Autodesk acessarem, o app precisa estar publicado em uma URL com HTTPS, por exemplo:

```txt
https://g5-gestao-coordenacao.onrender.com
```

## Caminho recomendado para esta fase

Para este app, o caminho mais simples é publicar como um serviço Node.js completo, porque ele tem:

- backend Express;
- frontend React;
- autenticação Autodesk;
- variáveis secretas no servidor.

Uma opção prática é usar **Render**, **Railway**, **Azure App Service** ou outro servidor Node.js com HTTPS.

## Variáveis de ambiente no servidor

No serviço publicado, configure:

```env
APS_CLIENT_ID=seu_client_id
APS_CLIENT_SECRET=seu_client_secret
APS_CALLBACK_URL=https://sua-url-publica.com/api/auth/callback
FRONTEND_URL=https://sua-url-publica.com
SESSION_SECRET=uma_frase_grande_e_segura
PORT=3000
```

No ambiente publicado, `APS_CALLBACK_URL` e `FRONTEND_URL` precisam usar a URL real do app, não `localhost`.

## Ajustar Callback URL no Autodesk Platform Services

No app APS **G5 Issues Planner**, adicione uma nova Callback URL:

```txt
https://sua-url-publica.com/api/auth/callback
```

Mantenha também a callback local se quiser continuar testando no seu computador:

```txt
http://localhost:3000/api/auth/callback
```

## Build e start

No servidor, os comandos esperados são:

```bash
npm install
npm run build
npm start
```

O `npm run build` gera o frontend em `dist/`.
O `npm start` roda o backend Express, que também entrega o frontend já compilado.

## Conectar ao ACC / Data Management

Depois que o app estiver publicado:

1. Acesse o Autodesk Construction Cloud Admin.
2. Entre como administradora da conta.
3. Vá para **Apps** ou **Custom Integrations**.
4. Adicione uma integração personalizada.
5. Informe o **Client ID** do app APS.
6. Use um nome claro, por exemplo:

```txt
G5 - Gestão e coordenação
```

7. Salve/ative a integração.

## Sobre aparecer “dentro” do Data Management

O cadastro como integração autoriza o app a ler/atualizar dados do ACC. Porém, a Autodesk pode não inserir um botão visual dentro de cada tela do Data Management para apps personalizados.

O resultado mais comum é:

- app autorizado no ACC/Admin;
- acesso ao app por uma URL própria;
- login Autodesk;
- leitura e atualização dos dados via APIs oficiais.

Se a Autodesk disponibilizar uma área de Apps/Integrações no seu tenant, o app pode aparecer ali como integração. Um botão embutido diretamente na tela de arquivos/projetos depende dos recursos que a Autodesk libera para a conta/produto.

## Próxima decisão

Para seguir, escolha onde publicar:

- Render: simples para começar.
- Azure App Service: mais corporativo.
- Servidor interno da G5: mais controle, mas exige infraestrutura.
