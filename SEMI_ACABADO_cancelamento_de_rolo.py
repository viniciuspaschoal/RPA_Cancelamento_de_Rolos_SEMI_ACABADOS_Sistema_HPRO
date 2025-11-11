# -*- coding: utf-8 -*-
"""
Automação HPro – Cancelamento de rolos para inventário
Autor: Lucas Texfyt

🧩 Dicas de segurança:
-------------------------------------------------------
Você pode configurar as credenciais de 2 formas:

1️⃣ Variáveis de ambiente (recomendado)
---------------------------------------
No Windows CMD:
    set HPRO_USER=vinicius_p
    set HPRO_PASS=Vp@tex29

⚠️ Essas variáveis duram apenas enquanto o terminal estiver aberto.
Para deixar permanente:
    Painel de Controle > Sistema > Configurações Avançadas > Variáveis de Ambiente

2️⃣ Input manual (automático no script)
---------------------------------------
Se as variáveis acima não existirem, o script vai perguntar:
    👤 Usuário HPro:
    🔒 Senha HPro:
-------------------------------------------------------
"""

import json
import os
import sys
import time
import getpass
from typing import List

import pyautogui
from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto.timings import TimeoutError as PywinautoTimeoutError

# ===========================#
# Configurações e constantes #
# ===========================#
APP_PATH = r"C:\Client HPro\Netrun.exe"
MAIN_TITLE = "Gerenciamento Administrativo"
LOGIN_WINDOW_TITLE = "Acesso ao Sistema"
RESTRICTED_ACCESS_TITLE = "Acesso restrito"
CANCEL_WINDOW_TITLE = "Cancelamento de rolo para inventário"
USER_CONFIRM_PANE = "Confirmação do usuário"
USER_ALERT_PANE = "Aviso ao usuário"

JSON_ROLOS_PATH = "rolos.json"

# ==============#
# Funções utilitárias
# ==============#
def load_rolos(path: str) -> List[str]:
    """Lê a lista de rolos do arquivo JSON"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        rolos = data.get("rolos", [])
        if not isinstance(rolos, list):
            raise ValueError("Campo 'rolos' precisa ser uma lista.")
        print(f"🔢 Rolos carregados: {rolos}")
        return [str(r) for r in rolos]
    except Exception as e:
        print(f"❌ Erro ao carregar rolos: {e}")
        sys.exit(1)


def safe_type_keys(ctrl, text: str, clear: bool = False):
    """Foca no campo e digita o texto com segurança"""
    ctrl.set_focus()
    time.sleep(0.2)
    if clear:
        ctrl.type_keys("^a{BACKSPACE}", pause=0.02)
        time.sleep(0.2)
    ctrl.type_keys(text, with_spaces=True, pause=0.02)


def click_first_button_with_text(root, text: str) -> bool:
    """Procura e clica no primeiro botão com o texto informado"""
    try:
        for b in root.descendants(control_type="Button"):
            if b.window_text() == text:
                b.click_input()
                return True
    except Exception:
        pass
    return False


# ==================#
# Fluxos principais #
# ==================#
def inicializar_hpro(user: str, password: str) -> Application:
    """Abre o HPro e faz o login inicial"""
    print("🚀 Iniciando HPro...")
    app = Application(backend="uia").start(APP_PATH)
    time.sleep(10)

    app = Application(backend="uia").connect(title=MAIN_TITLE)
    main_win = app.window(title=MAIN_TITLE)
    main_win.wait("visible", timeout=30)
    main_win.maximize()

    print("🔐 Procurando janela de login...")
    login_win = main_win.child_window(title=LOGIN_WINDOW_TITLE, control_type="Window")
    login_win.wait("visible", timeout=30)

    edits = login_win.descendants(control_type="Edit")
    print(f"🧭 Campos encontrados: {len(edits)}")

    if len(edits) < 2:
        raise RuntimeError("❌ Campos de login não encontrados!")

    safe_type_keys(edits[1], user, clear=True)
    time.sleep(0.5)
    safe_type_keys(edits[0], password, clear=True)
    time.sleep(0.5)

    login_win.child_window(title="Entrar", control_type="Button").click_input()
    print("✅ Login inicial concluído.")
    return app


def abrir_menu_cancelamento(main_win):
    """Abre o menu do cancelamento de rolos"""
    print("📂 Abrindo menu de cancelamento...")
    main_win.child_window(title="Etiquetas de Produção", control_type="MenuItem").click_input()
    time.sleep(0.8)
    main_win.child_window(title="Produtos em Elaboração", control_type="MenuItem").click_input()
    time.sleep(0.8)
    main_win.child_window(title="Processos Diversos", control_type="MenuItem").click_input()
    time.sleep(0.8)
    main_win.child_window(title="Cancelamento de rolos para inventário", control_type="MenuItem").click_input()
    time.sleep(1.5)

    pyautogui.click(x=755, y=342)  # clique do “dedinho”
    time.sleep(1)
    print("✅ Menu aberto e clique realizado.")


def login_restrito(main_win, user: str, password: str):
    """Login dentro do módulo de cancelamento"""
    print("🔐 Acesso restrito...")
    painel = main_win.child_window(title=RESTRICTED_ACCESS_TITLE, control_type="Window")
    painel.wait("visible", timeout=20)
    edits = painel.descendants(control_type="Edit")

    if len(edits) < 2:
        raise RuntimeError("❌ Campos de acesso restrito não encontrados!")

    safe_type_keys(edits[1], user, clear=True)
    time.sleep(0.2)
    safe_type_keys(edits[0], password, clear=True)
    painel.child_window(title="Acessar", control_type="Button").click_input()
    print("✅ Acesso restrito concluído.")


def processar_rolo(main_win, rolo: str):
    """Processa o cancelamento de um rolo"""
    print(f"\n=== 🔄 Iniciando rolo: {rolo} ===")

    inputs = main_win.descendants(control_type="Edit")
    if not inputs:
        inputs = main_win.descendants(control_type="Document") + main_win.descendants(control_type="Pane")

    if not inputs:
        raise RuntimeError("❌ Campo de código não encontrado!")

    campo_codigo = inputs[0]
    safe_type_keys(campo_codigo, rolo, clear=True)
    print(f"✅ Código {rolo} digitado.")

    cancelamento_win = main_win.child_window(title=CANCEL_WINDOW_TITLE, control_type="Window")
    cancelamento_win.wait("visible", timeout=10)

    cancelamento_win.child_window(title="Cancelar", control_type="Button").click_input()
    print("✅ Botão CANCELAR clicado.")
    time.sleep(1)

    confirm_win = main_win.child_window(title=USER_CONFIRM_PANE, control_type="Pane")
    aviso_win = main_win.child_window(title=USER_ALERT_PANE, control_type="Pane")

    if confirm_win.exists(timeout=3):
        print("🟢 Painel de confirmação detectado.")
        click_first_button_with_text(main_win, "Sim")
        time.sleep(0.6)
        click_first_button_with_text(main_win, "Sim")
        time.sleep(0.6)
        click_first_button_with_text(main_win, "OK")
        print(f"✅ Rolo {rolo} cancelado com sucesso.")

    elif aviso_win.exists(timeout=3):
        print("⚠️ Painel de aviso detectado (rolo já cancelado).")
        click_first_button_with_text(aviso_win, "OK")
    else:
        print("❌ Nenhum painel detectado – verifique o fluxo.")


def cancelamento_de_rolos(app, rolos, user, password):
    """Fluxo completo de cancelamento"""
    main_win = app.window(title=MAIN_TITLE)
    main_win.wait("visible", timeout=30)
    abrir_menu = True

    for rolo in rolos:
        if abrir_menu:
            abrir_menu_cancelamento(main_win)
            login_restrito(main_win, user, password)
            abrir_menu = False

        try:
            processar_rolo(main_win, rolo)
        except PywinautoTimeoutError:
            print(f"⏱️ Timeout ao processar {rolo}. Tentando reabrir menu...")
            abrir_menu = True
        except Exception as e:
            print(f"❌ Erro ao processar {rolo}: {e}")
            abrir_menu = True

    print("\n=== 🚀 TODOS OS ROLOS FORAM PROCESSADOS ===")


# ===== MAIN ===== #
def main():
    rolos = load_rolos(JSON_ROLOS_PATH)

    # 🧩 Tenta pegar do ambiente, senão pede input:
    user = os.environ.get("HPRO_USER") or input("👤 Usuário HPro: ")
    password = os.environ.get("HPRO_PASS") or getpass.getpass("🔒 Senha HPro: ")

    app = inicializar_hpro(user, password)
    time.sleep(3)
    cancelamento_de_rolos(app, rolos, user, password)


if __name__ == "__main__":
    main()
