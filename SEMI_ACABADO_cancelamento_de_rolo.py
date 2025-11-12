# -*- coding: utf-8 -*-
"""
Automação HPro – Cancelamento de rolos para inventário
Autor: Vinícius Paschoal
LinkedIn: https://www.linkedin.com/in/vinicius-paschoal/

🧩 Dicas de segurança:
-------------------------------------------------------
Você pode configurar as credenciais de 2 formas:

1️⃣ Variáveis de ambiente (recomendado)
---------------------------------------
No Windows CMD:
    set HPRO_USER=usuario
    set HPRO_PASS=senha

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


def find_dialog(main_win, title: str):
    """Tenta achar o diálogo por Window OU Pane; retorna o handler ou None."""
    # 1) tenta como Window
    try:
        dlg = main_win.child_window(title=title, control_type="Window")
        if dlg.exists(timeout=0.5):
            return dlg
    except Exception:
        pass
    # 2) tenta como Pane (algumas builds expõem assim)
    try:
        dlg = main_win.child_window(title=title, control_type="Pane")
        if dlg.exists(timeout=0.5):
            return dlg
    except Exception:
        pass
    return None


def processar_rolo(main_win, rolo: str) -> str:
    """
    Processa o cancelamento de um rolo.
    Retorna:
      - "sucesso"   -> se passou pelo fluxo de confirmação (2x "Sim" + "OK")
      - "aviso"     -> se apareceu popup de aviso/erro (OK) logo após Cancelar
      - "indefinido"-> se nenhum painel foi detectado
    """
    print(f"\n=== 🔄 Iniciando rolo: {rolo} ===")

    # Localiza/usa o primeiro campo editável
    inputs = main_win.descendants(control_type="Edit")
    if not inputs:
        inputs = main_win.descendants(control_type="Document") + main_win.descendants(control_type="Pane")
    if not inputs:
        raise RuntimeError("❌ Campo de código não encontrado!")

    campo_codigo = inputs[0]
    safe_type_keys(campo_codigo, rolo, clear=True)
    print(f"✅ Código {rolo} digitado.")

    # Garante a janela do módulo visível
    cancelamento_win = main_win.child_window(title=CANCEL_WINDOW_TITLE)
    cancelamento_win.wait("visible", timeout=10)

    # Clica CANCELAR
    cancelar_btn = cancelamento_win.child_window(title="Cancelar", control_type="Button")
    cancelar_btn.click_input()
    print("✅ Botão CANCELAR clicado.")

    # Espera reativamente por um diálogo de CONFIRMAÇÃO (Sim/Sim/OK) OU de AVISO (OK)
    t0 = time.time()
    timeout = 5.0  # segundos suficientes para aparecer o popup imediato
    dlg_tipo = "indefinido"

    while time.time() - t0 < timeout:
        # Tenta localizar os diálogos por título
        confirm_dlg = find_dialog(main_win, USER_CONFIRM_PANE)
        aviso_dlg   = find_dialog(main_win, USER_ALERT_PANE)

        if confirm_dlg:
            print("🟢 Painel de confirmação detectado.")
            click_first_button_with_text(main_win, "Sim")
            time.sleep(0.4)
            click_first_button_with_text(main_win, "Sim")
            time.sleep(0.4)
            click_first_button_with_text(main_win, "OK")
            print(f"✅ Rolo {rolo} cancelado com sucesso.")
            dlg_tipo = "sucesso"
            break

        if aviso_dlg:
            # Ex.: “Rolo já cancelado!” — OK e segue SEM reabrir menu/login
            print("⚠️ Painel de aviso detectado (ex.: rolo já cancelado).")
            # tenta no próprio diálogo; se não, tenta no main (fallback)
            if not click_first_button_with_text(aviso_dlg, "OK"):
                click_first_button_with_text(main_win, "OK")
            dlg_tipo = "aviso"
            break

        time.sleep(0.1)  # polling leve

    if dlg_tipo == "indefinido":
        print("❌ Nenhum painel detectado – verifique o fluxo.")

    return dlg_tipo


def cancelamento_de_rolos(app, rolos, user, password):
    """Fluxo completo de cancelamento conforme regra:
       - Se SUCESSO: reabrir menu + fazer login novamente para o próximo rolo.
       - Se AVISO/ERRO: NÃO reabrir; apenas colar o próximo código e seguir.
    """
    main_win = app.window(title=MAIN_TITLE)
    main_win.wait("visible", timeout=30)
    abrir_menu = True  # começa abrindo menu e fazendo login do módulo

    for rolo in rolos:
        if abrir_menu:
            abrir_menu_cancelamento(main_win)
            login_restrito(main_win, user, password)
            abrir_menu = False  # só volta a True em caso de SUCESSO

        try:
            status = processar_rolo(main_win, rolo)
            if status == "sucesso":
                # Regra pedida: após sucesso, reabre todo o caminho (menu + login)
                abrir_menu = True
            else:
                # Em aviso/erro/indefinido, mantém a tela e segue colando próximo código
                abrir_menu = False

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
