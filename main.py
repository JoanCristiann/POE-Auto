import time
import re

import pyautogui

# Tecla de emergência via WinAPI (F8) + fail-safe do PyAutoGUI
import ctypes
pyautogui.FAILSAFE = True  # mover o mouse para o canto sup. esquerdo aborta com exceção
VK_F8 = 0x77  # virtual key code F8

def emergency_pressed() -> bool:
    try:
        return bool(ctypes.windll.user32.GetAsyncKeyState(VK_F8) & 0x8000)
    except Exception:
        return False

# Acelera PyAutoGUI (remover pausas implícitas)
pyautogui.PAUSE = 0
pyautogui.MINIMUM_DURATION = 0
pyautogui.MINIMUM_SLEEP = 0

# Clipboard rápido
try:
    import pyperclip
except Exception:
    pyperclip = None


def get_clipboard_text(max_retries=3, sleep_between=0.03) -> str:
    txt = ""
    for _ in range(max_retries):
        try:
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(sleep_between)
            if pyperclip:
                t = pyperclip.paste()
            else:
                import subprocess
                t = subprocess.run(
                    ["powershell", "-NoProfile", "-Command", "Get-Clipboard"],
                    capture_output=True, text=True, timeout=1.5
                ).stdout
            t = (t or "").strip()
            if t:
                return t
        except Exception:
            pass
        time.sleep(sleep_between)
    return txt

# CONFIGURAÇÕES
CHAOS_POS = (550, 267)   # Chaos Orb
ITEM_POS  = (330, 450)   # Item a rolar
TARGETS = [
    "to Level of all Skill Gems",
    "to Maximum Life",
    # "r:\\+1\\s*to\\s*Level\\s*of\\s*all\\s*Skill\\s*Gems"  # exemplo regex
]
MAX_ROLLS = 1000
DELAY = 0.05  # teste entre 0.02 e 0.05 conforme estabilidade

MOUSE_ABORT_MARGIN = 6  # se o cursor for para perto (0,0), aborta


def corner_abort() -> bool:
    try:
        x, y = pyautogui.position()
        return x <= MOUSE_ABORT_MARGIN and y <= MOUSE_ABORT_MARGIN
    except Exception:
        return False


# Pequeno atraso para você focar a janela do jogo
time.sleep(1.0)

def postprocess_text(s: str) -> str:
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    fix = {
        "CHAOS RESlSTANCE": "CHAOS RESISTANCE",
        "RESlSTANCE": "RESISTANCE",
        "LEVel": "Level", "LEVeI": "Level", "LeVel": "Level",
        "SKlLL": "SKILL",
        "ATTRlBUTES": "ATTRIBUTES", "ATTRIBUTES": "Attributes",
        "CORRUPTED": "Corrupted",
    }
    for a, b in fix.items():
        s = s.replace(a, b)
    s = re.sub(r"\s*\+\s*", " +", s)
    s = re.sub(r"\s*%\s*", "%", s)
    s = re.sub(r"\s*to\s*", " to ", s, flags=re.I)
    # limpeza leve
    s = re.sub(r"\s*[\-–—•·]+\s*", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def extract_mod_lines(raw: str) -> str:
    blocks = [
        [ln.strip() for ln in blk.splitlines() if ln.strip()]
        for blk in raw.split('--------')
    ]
    mods = []
    for block in blocks:
        if not block:
            continue
        for ln in block:
            ln_l = ln.lower()
            if ln.startswith(('+', '-', 'Adds ')) or ('%' in ln) or ('increased' in ln_l) or ('reduced' in ln_l) or ('more ' in ln_l) or ('less ' in ln_l):
                if not ln_l.startswith(('item class', 'rarity', 'requirements', 'item level', 'sockets', 'note', 'quality')):
                    mods.append(ln)
    seen = set()
    unique_mods = []
    for m in mods:
        if m not in seen:
            unique_mods.append(m)
            seen.add(m)
    return "\n".join(unique_mods)



def text_matches_any_target(s: str) -> tuple[bool, str]:
    s_low = s.lower()
    for t in TARGETS:
        if t.startswith('r:'):
            pattern = t[2:]
            try:
                if re.search(pattern, s, flags=re.I):
                    return True, t
            except re.error:
                pass
        else:
            if t.lower() in s_low:
                return True, t
    return False, ""


# --------------------
# Loop principal
# --------------------

def main():
    print("Iniciando rolagem com Chaos Orb. Pressione F8 para abortar.")
    found = False
    all_rolls = []  # Lista para armazenar todos os rolls

    # 0) Fluxo por loop: ir no Chaos, botão direito, ir no item e aplicar

    try:
        for i in range(1, MAX_ROLLS + 1):
            # Checa tecla de emergência
            if emergency_pressed() or corner_abort():
                print("Abortado por tecla de emergência (F8) ou canto do mouse.")
                break

            # 1) Pegar Chaos Orb
            pyautogui.moveTo(CHAOS_POS[0], CHAOS_POS[1])
            time.sleep(DELAY)
            pyautogui.click(button='right')
            time.sleep(DELAY)

            # 2) Aplicar no item
            pyautogui.moveTo(ITEM_POS[0], ITEM_POS[1])
            time.sleep(DELAY)
            pyautogui.click(button='left')
            time.sleep(DELAY)

            # 3) Garantir tooltip do item para copiar
            pyautogui.moveTo(ITEM_POS[0], ITEM_POS[1])
            time.sleep(DELAY)
            raw = get_clipboard_text()

            # 4) Processar mods
            processed = postprocess_text(raw)
            mods_view = extract_mod_lines(raw)
            print(f"Roll {i:03d}\n{mods_view}\n")
            
            # Armazenar o roll na lista
            all_rolls.append({
                'roll_number': i,
                'mods': mods_view,
                'raw_text': raw,
                'processed_text': processed
            })

            matched, which = text_matches_any_target(processed)
            if matched:
                print(f"Alvo encontrado ({which}). Parando.")
                found = True
                break
    finally:
        # pausa breve para estabilizar
        time.sleep(0.02)

    if not found:
        print("Alvo não encontrado dentro do limite de rolagens ou abortado.")

    # Salvar todos os rolls em arquivo
    if all_rolls:
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        filename = f"chaos_rolls_{timestamp}.txt"

        with open(filename, 'w', encoding='utf-8') as f:
            f.write(f"Relatório de Chaos Orb Rolls - {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Total de rolls: {len(all_rolls)}\n")
            f.write(f"Alvo encontrado: {'Sim' if found else 'Não'}\n")
            f.write("=" * 50 + "\n\n")

            for roll_data in all_rolls:
                f.write(f"Roll {roll_data['roll_number']:03d}\n")
                f.write("-" * 20 + "\n")
                f.write(f"{roll_data['mods']}\n\n")

        print(f"Arquivo salvo: {filename}")


if __name__ == "__main__":
    main()