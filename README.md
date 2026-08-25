# Gerador de Relatório Semanal — versão com telinha (Streamlit)

Versão do `generate_report.py` original com uma interface web local, para
quem não quer/sabe rodar comando por terminal.

## Estrutura

```
report_app/
├── app.py             # interface Streamlit (rode este arquivo)
├── report_core.py      # toda a lógica de leitura/geração (igual ao script original)
├── templates/
│   └── email_template.html.j2
├── requirements.txt
└── README.md
```

## Instalação (primeira vez)

```bash
pip install -r requirements.txt --break-system-packages
```

## Uso

```bash
streamlit run app.py
```

Isso abre uma aba no navegador (geralmente `http://localhost:8501`). Lá:

1. Suba o arquivo `.xlsx` da semana.
2. Escolha, no menu, qual aba é a da semana (ex: `31-08 a 04-09`) — o app já
   esconde as abas `Config` e `KPIs` dessa lista, porque elas têm uso fixo.
3. Clique em **Gerar relatório**.
4. Veja a prévia na tela e clique em **Baixar HTML**.
5. Abra o HTML baixado no navegador, selecione tudo (Ctrl+A) e cole no corpo
   do email (Gmail/Outlook) — não cole o código-fonte.

## O que mudou em relação à versão de linha de comando

- Não precisa mais rodar `python3 generate_report.py ... --sheet "..."` —
  é tudo pela tela.
- A aba de dados agora é escolhida num menu, em vez de digitada.
- A aba **KPIs** continua sendo lida automaticamente pelo nome fixo `KPIs`
  (atualize os números na planilha antes de subir o arquivo).
- **Período e Destaques continuam sem campo editável nesta primeira
  versão** — vêm da aba `Config` da planilha, se existir, ou são calculados
  automaticamente (igual à versão CLI). Isso pode virar um próximo passo,
  se fizer sentido.

## Observação

A lógica de leitura da planilha (agrupamento, cores de tag/status, KPIs,
etc.) é exatamente a mesma do `generate_report.py` original — só foi movida
para `report_core.py` para ser reaproveitada pela interface. Nenhuma regra
de negócio foi alterada.

## Hospedando no Streamlit Community Cloud (repositório privado)

Passo a passo pra deixar o app acessível por um link, restrito só a quem
você convidar (por causa dos dados sensíveis da planilha):

1. **Crie um repositório privado no GitHub** e suba os arquivos desta pasta
   (`app.py`, `report_core.py`, `templates/`, `requirements.txt`,
   `.gitignore`). O `.gitignore` já impede que qualquer `.xlsx` de teste
   suba junto por engano.

   ```bash
   cd report_app
   git init
   git add .
   git commit -m "Primeira versão do app"
   git branch -M main
   git remote add origin https://github.com/SEU_USUARIO/SEU_REPO.git
   git push -u origin main
   ```

   (Crie o repositório como **privado** na tela do GitHub antes do `git push`.)

2. **Acesse [share.streamlit.io](https://share.streamlit.io)** e faça login
   com a sua conta GitHub. Autorize o Streamlit a acessar seus repositórios
   privados quando ele pedir.

3. Clique em **"New app"**, escolha o repositório, a branch (`main`) e o
   arquivo principal (`app.py`). Clique em **Deploy**.

4. Como o repositório é privado, o app nasce **privado automaticamente** —
   ninguém acessa sem permissão, mesmo tendo o link.

5. Pra dar acesso a alguém da equipe: abra o app publicado, clique em
   **"Share"** no canto superior direito e adicione o **email** da pessoa.
   Ela recebe um convite e consegue entrar com login do Google (se o email
   for do Google) ou por um link de acesso único.

**Limitação do plano gratuito:** só é permitido **um app privado por vez**.
Se um dia precisar publicar outro app privado, é preciso tornar este
público ou removê-lo antes.

**Sobre os dados:** o app não guarda a planilha em lugar nenhum — cada
pessoa sobe seu próprio `.xlsx` na hora de gerar o relatório, e o arquivo
não fica salvo entre sessões. Ainda assim, os dados passam pela
infraestrutura do Streamlit durante o uso (tudo via HTTPS). Se isso for
uma restrição para a instituição, a alternativa é hospedar num servidor
próprio em vez do Streamlit Community Cloud.
