CONSOLIDADOR EFD ICMS — PUBLICAÇÃO NO RENDER

OBJETIVO
Publicar o sistema em um endereço fixo HTTPS, independente do computador local.

ARQUIVOS DO PACOTE
- app.py: aplicação Flask
- requirements.txt: dependências
- render.yaml: configuração automática do serviço
- .gitignore: arquivos que não devem ir para o GitHub

PASSO 1 — CRIAR UM REPOSITÓRIO NO GITHUB
1. Entre no GitHub e crie um repositório, por exemplo: consolidador-efd-icms.
2. Envie os quatro arquivos deste pacote para a raiz do repositório.

PASSO 2 — PUBLICAR NO RENDER
1. Entre no Render usando sua conta GitHub.
2. Clique em New > Blueprint.
3. Selecione o repositório consolidador-efd-icms.
4. Confirme a criação do serviço descrito em render.yaml.
5. Aguarde o deploy terminar.
6. O Render exibirá um link fixo semelhante a:
   https://consolidador-efd-icms.onrender.com

PROTEÇÃO POR SENHA (RECOMENDADA)
No painel do serviço, abra Environment e crie:
- APP_USERNAME = efd
- APP_PASSWORD = escolha-uma-senha-forte

Depois salve. O navegador passará a solicitar usuário e senha.

SOBRE O PLANO GRATUITO
O endereço é fixo, mas o serviço gratuito adormece após um período sem uso.
No primeiro acesso depois disso, pode levar aproximadamente um minuto para abrir.
Para uso contínuo, altere o plano do serviço para uma instância paga que não adormeça.

TESTE DE SAÚDE
Abra:
https://SEU-LINK.onrender.com/health
A resposta deve ser: ok
