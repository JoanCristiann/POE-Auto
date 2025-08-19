import time
import re
import os
import pyautogui
import pytesseract
import cv2
import numpy as np
from PIL import Image

# Tecla de emergência via WinAPI (F8) + fail-safe do PyAutoGUI
import ctypes
pyautogui.FAILSAFE = True  # mover o mouse para o canto sup. esquerdo aborta com exceção
VK_F8 = 0x77  # virtual key code F8

def emergency_pressed() -> bool:
    try:
        return bool(ctypes.windll.user32.GetAsyncKeyState(VK_F8) & 0x8000)
    except Exception:
        return False

# Se precisar, defina o executável do tesseract:
# pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# CONFIGURAÇÕES
TOOLTIP_REGION = (84, 198, 474, 274)
CHAOS_POS = (550, 267)   # Chaos Orb
ITEM_POS  = (330, 450)   # Item a rolar
TARGETS = [
    "+1 to Level of all Skill Gems",
    # "r:\\+1\\s*to\\s*Level\\s*of\\s*all\\s*Skill\\s*Gems"  # exemplo regex
]
MAX_ROLLS = 300
ROLL_DELAY = 0.01  # mais rápido; ajuste se necessário
EMERGENCY_KEY = 'f8'
USE_AUTOCROP = True
def release_shift():
    try:
        pyautogui.keyUp('shift')
    except Exception:
        pass

MOUSE_ABORT_MARGIN = 6  # se o cursor for para perto (0,0), aborta

def corner_abort() -> bool:
    try:
        x, y = pyautogui.position()
        return x <= MOUSE_ABORT_MARGIN and y <= MOUSE_ABORT_MARGIN
    except Exception:
        return False


# Pequeno atraso para você focar a janela do jogo
time.sleep(1.0)

# --------------------
# Pré-processamento / OCR
# --------------------

def upscale(img, scale=1.8):
    h, w = img.shape[:2]
    return cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC)


def unsharp(gray, amount=1.0, radius=3):
    blur = cv2.GaussianBlur(gray, (0, 0), radius)
    sharp = cv2.addWeighted(gray, 1 + amount, blur, -amount, 0)
    return np.clip(sharp, 0, 255).astype(np.uint8)


def to_black_text(th):
    # Garantir fundo branco e texto preto
    return cv2.bitwise_not(th) if th.mean() > 127 else th


def prepare_candidates(pil_img):
    rgb = np.array(pil_img.convert("RGB"))
    hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV)
    V = hsv[:, :, 2]
    bgr = cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)
    B = bgr[:, :, 0]

    cands = []

    # Candidate 1: V channel + gamma + unsharp + Otsu
    v1 = upscale(V, 1.8)
    v1 = np.clip(((v1 / 255.0) ** (1 / 1.2)) * 255, 0, 255).astype(np.uint8)  # gamma leve
    v1 = unsharp(v1, 0.8, 2)
    _, th1 = cv2.threshold(v1, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    th1 = to_black_text(th1)
    cands.append(th1)

    # Candidate 2: V channel + adaptive
    v2 = upscale(V, 1.8)
    v2 = cv2.bilateralFilter(v2, 5, 35, 35)
    th2 = cv2.adaptiveThreshold(v2, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                cv2.THRESH_BINARY, 31, 10)
    th2 = to_black_text(th2)
    cands.append(th2)

    # Candidate 3: Blue channel (texto azul) + Otsu
    b1 = upscale(B, 1.8)
    b1 = unsharp(b1, 0.8, 2)
    _, th3 = cv2.threshold(b1, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    th3 = to_black_text(th3)
    cands.append(th3)

    # Padding
    cands = [cv2.copyMakeBorder(c, 8, 8, 8, 8, cv2.BORDER_CONSTANT, value=255) for c in cands]
    return cands


def try_autocrop_tooltip(bin_img):
    if bin_img is None or len(bin_img.shape) != 2 or bin_img.size == 0:
        return bin_img
    inv = cv2.bitwise_not(bin_img)
    kernel = np.ones((3, 3), np.uint8)
    closed = cv2.morphologyEx(inv, cv2.MORPH_CLOSE, kernel, iterations=1)
    contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return bin_img
    H, W = bin_img.shape
    best_rect, best_area = None, 0
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        area = w * h
        if 0.02 * W * H <= area <= 0.98 * W * H and w > 60 and h > 60:
            if area > best_area:
                best_area = area
                best_rect = (x, y, w, h)
    if best_rect is None:
        return bin_img
    x, y, w, h = best_rect
    crop = bin_img[y:y + h, x:x + w]
    crop = cv2.copyMakeBorder(crop, 6, 6, 6, 6, cv2.BORDER_CONSTANT, value=255)
    return crop


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


def ocr_with_conf(img_bin):
    config = "--oem 1 --psm 4 -c preserve_interword_spaces=1"
    data = pytesseract.image_to_data(img_bin, lang="eng", config=config, output_type=pytesseract.Output.DICT)
    confs = [int(c) for c in data["conf"] if c != "-1"]
    conf_avg = sum(confs) / len(confs) if confs else 0.0

    # Reconstrói por linhas
    lines = {}
    n = len(data["text"])
    for i in range(n):
        word = data["text"][i].strip()
        if not word or data["conf"][i] == "-1":
            continue
        key = (data.get("block_num", [0])[i], data.get("par_num", [0])[i], data.get("line_num", [0])[i])
        lines.setdefault(key, []).append(word)

    ordered = sorted(lines.keys())
    text_lines = [" ".join(lines[k]) for k in ordered]
    text = "\n".join(text_lines)
    return conf_avg, text


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

def main():
    print("Iniciando rolagem com Chaos Orb. Pressione F8 para abortar.")
    found = False

    # Seleciona o Chaos uma vez (botão direito) e mantém SHIFT pressionado
    pyautogui.moveTo(CHAOS_POS[0], CHAOS_POS[1])
    time.sleep(0.05)
    pyautogui.click(button='right')  # seleciona chaos
    time.sleep(0.05)
    pyautogui.keyDown('shift')       # mantém selecionado

    try:
        for i in range(1, MAX_ROLLS + 1):
            # Checa tecla de emergência
            if emergency_pressed() or corner_abort():
                print("Abortado por tecla de emergência (F8) ou canto do mouse.")
                break

            # 1) Aplicar Chaos no item (não volta ao Chaos)
            pyautogui.moveTo(ITEM_POS[0], ITEM_POS[1])
            pyautogui.click(button='left')

            time.sleep(ROLL_DELAY)

            # 2) Captura tooltip e OCR
            pil_img = pyautogui.screenshot(region=TOOLTIP_REGION)
            cands = prepare_candidates(pil_img)
            if USE_AUTOCROP:
                cands = [try_autocrop_tooltip(c) for c in cands]

            best_conf, best_text = -1, ""
            for c in cands:
                conf, txt = ocr_with_conf(c)
                if conf > best_conf:
                    best_conf, best_text = conf, txt

            processed = postprocess_text(best_text).replace('\n', ' ')
            print(f"Roll {i:03d} | conf={best_conf:.1f} | {processed}")

            # 3) Verifica alvos
            matched, which = text_matches_any_target(processed)
            if matched:
                print(f"Alvo encontrado ({which}). Parando.")
                found = True
                break
    finally:
        # Solta SHIFT mesmo em caso de abort/erro
        release_shift()
        time.sleep(0.05)

    if not found:
        print("Alvo não encontrado dentro do limite de rolagens ou abortado.")


if __name__ == "__main__":
    main()