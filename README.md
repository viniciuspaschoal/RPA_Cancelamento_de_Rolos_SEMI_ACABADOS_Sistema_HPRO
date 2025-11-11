# 🤖 RPA: Cancelamento de Rolos Semi-Acabados (HPro)

![Status](https://img.shields.io/badge/Status-Em%20Desenvolvimento-blue)

Este projeto é uma automação em **Python** desenvolvida para realizar o **cancelamento automático de rolos** no sistema **HPro (Gerenciamento Administrativo)**.

O objetivo é **substituir tarefas manuais repetitivas** no processo de inventário por uma rotina automática confiável e segura, reduzindo erros e acelerando o fluxo de trabalho da equipe.

## ✨ Tecnologias Utilizadas

* **Python**
* **pywinauto**
* **pyautogui**

---

## ⚙️ Como Funciona

A automação utiliza as bibliotecas **pywinauto** e **pyautogui** para interagir com as janelas do sistema HPro, simulando as ações de um operador humano.

O script executa as seguintes etapas:

1.  Inicia o aplicativo `Netrun.exe` (cliente HPro).
2.  Realiza o login automático.
3.  Acessa o menu **"Cancelamento de rolos para inventário"**.
4.  Lê uma lista de rolos do arquivo `rolos.json`.
5.  Cancela cada rolo um por um, tratando confirmações e avisos.
6.  Exibe logs no terminal com o status de cada operação.

---

## 🔐 Gerenciamento de Credenciais

O script foi projetado para **nunca armazenar senhas diretamente no código** (`.py`), garantindo mais segurança durante a execução.

Existem **duas formas** de fornecer as credenciais:

### 1. Variáveis de Ambiente (Modo Automático)

Você pode definir as credenciais antes de rodar o script, usando o terminal. Essas variáveis ficam ativas apenas enquanto o terminal estiver aberto.

**No Windows (CMD/PowerShell):**
```
set HPRO_USER=seu_usuario
set HPRO_PASS=sua_senha
```

O Python as lê automaticamente usando:
```
usuario = os.environ.get("HPRO_USER")
senha = os.environ.get("HPRO_PASS")
```

Dessa forma, a senha não aparece em nenhum arquivo.


### 2. Input Manual (Modo Interativo)

Se as variáveis de ambiente não existirem, o script pedirá o login na tela:

👤 Usuário HPro: seu_usuario
🔒 Senha HPro: 

Durante a digitação da senha, nada é exibido (nem asteriscos). Esse comportamento é proposital e vem do módulo getpass, garantindo que a senha não apareça visualmente no terminal.
