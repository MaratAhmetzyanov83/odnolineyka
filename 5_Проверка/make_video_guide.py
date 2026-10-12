# -*- coding: utf-8 -*-
"""Видеогайд «Однолинейка v3.12 — как пользоваться»: слайды со снимками листов модели объекта + озвучка (голос Windows).

    py make_video_guide.py [папка модели] [выход.mp4]
       по умолчанию: C:/Users/marat/Desktop/Модель_ЖД_2ввода  ->  1_Таблицы_Excel/Видеогайд_Однолинейка_v3.12.mp4

Снимки листов — свой скрытый Excel (экспорт областей в PDF -> PNG через PyMuPDF), ПЗ — из ПЗ_<папка>.docx (Word -> PDF).
Озвучка — System.Speech (голос Microsoft Irina), сборка — ffmpeg. Excel / Word — свои экземпляры, закрываются только они.
Модель пересобрать: py 5_Проверка/build_demo_object.py "<папка модели>".
"""
import io
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import wave

import fitz                                  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
DEF_MODEL = r"C:\Users\marat\Desktop\Модель_ЖД_2ввода"
DEF_OUT = os.path.join(ROOT, "1_Таблицы_Excel", "Видеогайд_Однолинейка_v3.12.mp4")
W, H = 1920, 1080
FONT = r"C:\Windows\Fonts\segoeui.ttf"
FONT_B = r"C:\Windows\Fonts\segoeuib.ttf"
BLUE, DARK, LIGHT = (47, 85, 151), (31, 41, 55), (243, 246, 250)

# (заголовок, снимок или None, текст на экране, текст озвучки)
SLIDES = [
    ("Однолинейка v3.12", None,
     ["Щит — из таблицы Excel в схему AutoCAD", "Как пользоваться таблицей и надстройкой"],
     "Однолинейка, версия три двенадцать. Как пользоваться таблицей щита и надстройкой на ленте Excel."),
    ("Что в комплекте", None,
     ["Шаблон таблицы: 1_Таблицы_Excel / Однолинейка_шаблон.xlsx",
      "Надстройка Excel: 3_Макросы_Excel / Однолинейка.xlam — вкладка «Однолинейка»",
      "AutoCAD: DLN — длины с планов, SXDRAW — схема, CPDRAW — компоновка, CableJournal — кабельный журнал",
      "Один щит — одна книга. Книги объекта — в одной папке (можно с подпапками)"],
     "В комплекте шаблон таблицы, надстройка для Excel и программы для AutoCAD. Правило простое: один щит — одна книга. "
     "Книги объекта лежат в одной папке, можно с подпапками."),
    ("Пульт — первый лист", "pult",
     ["Порядок работы со ссылками на листы", "Ошибки «Однолинейки» и проверка КЗ", "Итоги и ссылки на все листы"],
     "Первый лист — пульт. Здесь порядок работы со ссылками на листы, счётчик ошибок, проверка токов короткого замыкания и итоги щита."),
    ("1. Исходные данные", "src",
     ["Строка = линия: номер (R.1.01.1 — код, этаж, помещение, №), мощность, длина",
      "Длины, гофру и лоток заполняет DLN с планов AutoCAD",
      "Кабель, сечение, автомат, ΔU, ток КЗ — считаются сами",
      "Код в номере линии берётся из «Справочной»"],
     "Шаг первый — исходные данные. Одна строка — одна линия: номер, мощность и длина. Длины с планов заполняет программа DLN. "
     "Кабель, сечение, автомат, потери напряжения и ток короткого замыкания таблица считает сама по коду линии."),
    ("Кнопки столбцов", "ribbon",
     ["Только ввод — видны только столбцы, которые вы заполняете",
      "Ввод + расчёт — плюс ток, автомат, кабель, ΔU, КЗ, проверки",
      "Все столбцы — со служебными расчётами"],
     "Чтобы не путаться в восьмидесяти столбцах, на ленте есть кнопки: только ввод, ввод плюс расчёт и все столбцы. Данные они не меняют."),
    ("2. Разбивка по щитам", "rz",
     ["Для каждой строки — щит и номер линии", "Порядок строк = порядок мест на схеме"],
     "Шаг второй — разбивка по щитам. Для каждой строки укажите щит и номер линии. Порядок строк — это порядок мест на схеме."),
    ("3. Однолинейка", "odn",
     ["Белые столбцы — ввод: шина, автомат «авто» или текст, 30 мА, аппараты, модуль, канал, клемма",
      "Столбец AE — статус строки: что проверить",
      "Двигатели — автомат D, проверка пускового тока"],
     "Шаг третий — лист однолинейка. Белые столбцы — ввод: шина, автомат «авто» или свой текст, тридцать миллиампер, аппараты, "
     "модуль, канал и клемма. Столбец статуса подсказывает, что проверить. Для двигателей таблица выбирает автомат с характеристикой D."),
    ("Справочная — коды потребителей", "ref",
     ["Код → наименование, кабель, напряжение, фазы, группа Кс, УГО",
      "Новый код — новая строка. Щ — щит, ЛФ — лифт, НС — насос, LM — освещение МОП",
      "Каждый блок подписан: что это и где используется"],
     "Справочная хранит коды потребителей: наименование, кабель, напряжение, группу коэффициента спроса и значок на схеме. "
     "Новый код — это новая строка. Каждый блок подписан."),
    ("4. Нагрузка щита", "nag",
     ["Руст, Кс по СП 256, Рр, cos φ, Iр по наиболее нагруженной фазе",
      "Секция 1 / секция 2 / весь щит (аварийный режим)",
      "Ко = 0,9 — только если в щите есть бытовая нагрузка"],
     "Шаг четвёртый — нагрузка щита. Таблица считает установленную и расчётную мощность по своду правил двести пятьдесят шесть, "
     "ток по наиболее нагруженной фазе, а при двух вводах — по секциям и в аварийном режиме."),
    ("Сеть щита: питание, нижестоящие, токи КЗ", "net",
     ["Раздел А — откуда питается щит: ΔU, Zп, Iкз на шинах, вводы 2 и 3",
      "Раздел Б — кого питает щит; раздел В — токи КЗ по линиям для печати",
      "Icu аппаратов щита — проверка отключающей способности"],
     "Лист сеть щита. Раздел А — откуда питается щит: потери напряжения и ток короткого замыкания на шинах, второй и третий ввод. "
     "Раздел Б — нижестоящие щиты. Раздел В — токи короткого замыкания по линиям для печати."),
    ("Связать с вышестоящим", "link",
     ["В книге нижестоящего щита: кнопка → файл вышестоящего → линия-фидер → ввод",
      "Связи встают в обе книги: питание сверху, нагрузка снизу",
      "Передаётся номер автомата фидера (QF1), автомат, кабель и длина — для схемы"],
     "Щиты связывает кнопка «Связать с вышестоящим». В книге нижестоящего щита выберите файл вышестоящего и линию-фидер. "
     "Связи появятся в обеих книгах. Вместе с питанием приходят номер автомата фидера, кабель и длина — они нужны для схемы."),
    ("Два ввода и селективность", "inputs",
     ["Каждый ввод — свой фидер, ΔU и Iкз; с АВР — худший из вводов",
      "Фидер несёт свою секцию, ток — аварийный (весь щит)",
      "Ввод ≥ 1,6 × отходящих; иначе — выключатель нагрузки QS (QS1, QS2, QF3…)"],
     "Щит с двумя вводами подключается двумя фидерами. Каждый несёт свою секцию, без задвоения нагрузки. "
     "Вводной аппарат подбирается по селективности, а если автомат невозможен, ставится выключатель нагрузки QS."),
    ("Дерево щитов", "tree",
     ["Все щиты объекта: кто от кого питается", "Руст, Iр, ΔU и Iкз на шинах, проверки",
      "Запуск из любой книги — сам находит главный щит и подпапки"],
     "Кнопка «Дерево щитов» показывает все щиты объекта: кто от кого питается, нагрузки, потери и токи короткого замыкания. "
     "Её можно запускать из любой книги объекта."),
    ("ПЗ по объекту", "pz",
     ["Кнопка «ПЗ по объекту» → ПЗ_объект.xlsx и ПЗ_объект.docx",
      "Нагрузки щитов, режимы, категории, ΔU и КЗ, дерево, замечания",
      "Расчётная часть ПЗ в Word — с таблицами и выводами"],
     "Для стадии П нажмите «ПЗ по объекту». Рядом появятся два файла: таблицы в Excel и расчётная часть пояснительной записки в Word."),
    ("Спецификация, бирки, наклейки", None,
     ["Обновить спецификацию → изменения подсвечиваются; Принять изменения → ревизия +1",
      "Бирки и наклейки под аппараты — кнопками, листы появляются сами",
      "Позиции — в листе «Каталог»"],
     "Спецификация обновляется кнопкой, изменения для снабжения подсвечиваются. Бирки и наклейки под аппараты тоже создаются кнопками."),
    ("AutoCAD", None,
     ["DLN — длины трасс с планов в «Исходные данные»",
      "SXDRAW — однолинейная схема по листу «Однолинейка», SXUPDATE — обновить",
      "CPDRAW — компоновка щита; CableJournal — кабельный журнал"],
     "В AutoCAD: DLN переносит длины с планов в таблицу, SXDRAW рисует однолинейную схему, CPDRAW — компоновку щита, "
     "CableJournal — кабельный журнал."),
    ("Старые книги и два ПК", None,
     ["Обновить старую книгу: py 5_Проверка/update_book.py «книга.xlsx»",
      "Начали работу — Получить.cmd; закончили — Отправить.cmd",
      "Что сделано и что дальше — файл СОСТОЯНИЕ.md"],
     "Старую книгу объекта обновляет одна команда update book. На двух компьютерах: начали работу — «Получить», закончили — «Отправить». "
     "Что сделано и что дальше — в файле «Состояние»."),
]

# снимки: (лист, область) из книг модели
SHOTS = {
    "pult": ("ГРЩ.xlsx", "Пульт", "A1:D30"),
    "src": ("ЩК-1.xlsx", "Исходные данные", "A1:V36"),
    "rz": ("ЩК-1.xlsx", "Разбивка_по_щитам", "A1:Q16"),
    "odn": ("ЩС-1.xlsx", "Однолинейка", "A1:AE12"),
    "ref": ("ГРЩ.xlsx", "Справочная", "P1:AB40"),
    "nag": ("ЩС-1.xlsx", "Нагрузка_щитов", "R4:AH12"),
    "net": ("ЩР-1.xlsx", "Сеть_щита", "B1:L41"),
    "link": ("ЩР-1.xlsx", "Сеть_щита", "B4:L8"),
    "inputs": ("ЩС-1.xlsx", "Сеть_щита", "U4:AP11"),
    "tree": ("ГРЩ.xlsx", "Дерево_щитов", "A1:H11"),
}


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def kill(app, pid):
    try:
        app.Quit()
    except Exception:
        pass
    try:
        os.kill(pid, signal.SIGTERM)
    except Exception:
        pass


def pdf_to_png(pdf, png, zoom=2.5):
    doc = fitz.open(pdf)
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    pix.save(png)
    doc.close()
    im = Image.open(png).convert("RGB")
    bg = Image.new("RGB", im.size, (255, 255, 255))
    from PIL import ImageChops
    box = ImageChops.difference(im, bg).getbbox()
    if box:
        im = im.crop((max(0, box[0] - 12), max(0, box[1] - 12), min(im.width, box[2] + 12), min(im.height, box[3] + 12)))
    im.save(png)


def take_shots(model, tmp):
    app, pid = excel()
    out = {}
    try:
        cache = {}
        for key, (book, sheet, rng) in SHOTS.items():
            path = None
            for r, _, fs in os.walk(model):
                if book in fs:
                    path = os.path.join(r, book)
                    break
            if not path:
                print("  нет книги", book)
                continue
            wb = cache.get(path) or app.Workbooks.Open(path, 0, True)
            cache[path] = wb
            ws = wb.Worksheets(sheet)
            if sheet in ("Исходные данные", "Однолинейка"):
                try:
                    if ws.ProtectContents:
                        ws.Unprotect("sx")
                    ws.Outline.ShowLevels(0, 1)          # уровень «Только ввод»
                except Exception:
                    pass
            if sheet == "Исходные данные":
                ws.Rows("3:20").Hidden = True               # служебные пустые строки шаблона — не на снимок (книга не сохраняется)
            ps = ws.PageSetup
            ps.PrintArea = rng
            ps.Orientation = 2
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            pdf = os.path.join(tmp, key + ".pdf")
            ws.ExportAsFixedFormat(0, pdf)
            png = os.path.join(tmp, key + ".png")
            pdf_to_png(pdf, png)
            out[key] = png
        for wb in cache.values():
            wb.Close(False)
    finally:
        kill(app, pid)
    # ПЗ: первая страница Word
    docx = [os.path.join(model, f) for f in os.listdir(model) if f.startswith("ПЗ_") and f.endswith(".docx")]
    if docx:
        import win32com.client
        wd = win32com.client.DispatchEx("Word.Application")
        try:
            d = wd.Documents.Open(docx[0], False, True)
            pdf = os.path.join(tmp, "pz.pdf")
            d.ExportAsFixedFormat(pdf, 17)
            d.Close(False)
        finally:
            wd.Quit()
        png = os.path.join(tmp, "pz.png")
        pdf_to_png(pdf, png, 2.0)
        out["pz"] = png
    out["ribbon"] = ribbon_png(tmp)
    return out


def ribbon_png(tmp):
    """схематичная вкладка «Однолинейка» на ленте"""
    groups = [("Спецификация", ["Обновить\nспецификацию", "Принять\nизменения"]),
              ("Печать", ["Бирки", "Наклейки"]),
              ("Связи щитов", ["Связать с\nвышестоящим", "Обновить\nсвязи", "Дерево\nщитов"]),
              ("Стадия П", ["ПЗ по\nобъекту"]),
              ("Столбцы", ["Только\nввод", "Ввод +\nрасчёт", "Все\nстолбцы"]),
              ("Справка", ["Как\nпользоваться"])]
    f = ImageFont.truetype(FONT, 22)
    fb = ImageFont.truetype(FONT_B, 22)
    bw, gap = 150, 18
    width = sum(len(b) * bw + gap * 2 for _, b in groups) + 40
    im = Image.new("RGB", (width, 230), (250, 250, 250))
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, width, 40), fill=(43, 87, 154))
    d.text((20, 6), "Файл   Главная   Вставка   …   Однолинейка", font=fb, fill="white")
    x = 20
    for name, btns in groups:
        x0 = x
        for b in btns:
            d.rounded_rectangle((x + 6, 56, x + bw - 6, 170), 10, fill=(225, 234, 247), outline=(160, 180, 210))
            d.multiline_text((x + bw / 2, 113), b, font=f, fill=DARK, anchor="mm", align="center")
            x += bw
        d.text(((x0 + x) / 2, 196), name, font=f, fill=(90, 90, 90), anchor="mm")
        x += gap
        d.line((x, 52, x, 215), fill=(200, 200, 200), width=2)
        x += gap
    p = os.path.join(tmp, "ribbon.png")
    im.save(p)
    return p


def slide(n, title, shot, lines, path):
    im = Image.new("RGB", (W, H), LIGHT)
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, W, 110), fill=BLUE)
    d.text((60, 55), title, font=ImageFont.truetype(FONT_B, 56), fill="white", anchor="lm")
    d.text((W - 60, 55), f"{n}/{len(SLIDES)}", font=ImageFont.truetype(FONT, 34), fill=(220, 230, 245), anchor="rm")
    ft = ImageFont.truetype(FONT, 38 if shot else 46)
    if shot and os.path.exists(shot):
        txt_h = 70 + 58 * len(lines)
        area = (60, 140, W - 60, H - txt_h - 40)
        sh = Image.open(shot)
        k = min((area[2] - area[0]) / sh.width, (area[3] - area[1]) / sh.height, 1.6)
        sh = sh.resize((int(sh.width * k), int(sh.height * k)), Image.LANCZOS)
        x = (W - sh.width) // 2
        y = area[1] + (area[3] - area[1] - sh.height) // 2
        d.rectangle((x - 4, y - 4, x + sh.width + 4, y + sh.height + 4), fill=(200, 205, 215))
        im.paste(sh, (x, y))
        y = H - txt_h
        d.rectangle((0, y - 20, W, H), fill=DARK)
        for i, t in enumerate(lines):
            d.text((80, y + 10 + 58 * i), "• " + t, font=ft, fill="white")
    else:
        y = 300 if n > 1 else 400
        if n == 1:
            d.text((W // 2, 330), "ОДНОЛИНЕЙКА", font=ImageFont.truetype(FONT_B, 120), fill=BLUE, anchor="mm")
        for i, t in enumerate(lines):
            if n == 1:
                d.text((W // 2, 520 + 80 * i), t, font=ft, fill=DARK, anchor="mm")
            else:
                d.text((120, y + 110 * i), "• " + t, font=ft, fill=DARK)
    im.save(path)


def speak(text, wav):
    ps = ("Add-Type -AssemblyName System.Speech; $s=New-Object System.Speech.Synthesis.SpeechSynthesizer; "
          "$s.SelectVoice('Microsoft Irina Desktop'); $s.Rate=1; "
          f"$s.SetOutputToWaveFile('{wav}'); $s.Speak([IO.File]::ReadAllText('{wav}.txt',[Text.Encoding]::UTF8)); $s.Dispose()")
    with io.open(wav + ".txt", "w", encoding="utf-8") as f:
        f.write(text)
    subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True, capture_output=True)
    with wave.open(wav) as w:
        return w.getnframes() / w.getframerate()


def run(model, out):
    ff = shutil.which("ffmpeg")
    if not ff:
        raise SystemExit("нет ffmpeg")
    tmp = tempfile.mkdtemp(prefix="sx_video_")
    print("снимки листов…")
    shots = take_shots(model, tmp)
    segs = []
    for i, (title, key, lines, voice) in enumerate(SLIDES, 1):
        png = os.path.join(tmp, f"s{i:02d}.png")
        slide(i, title, shots.get(key) if key else None, lines, png)
        wav = os.path.join(tmp, f"s{i:02d}.wav")
        dur = speak(voice, wav) + 0.8
        seg = os.path.join(tmp, f"s{i:02d}.mp4")
        subprocess.run([ff, "-y", "-loglevel", "error", "-loop", "1", "-i", png, "-i", wav,
                        "-af", "apad=pad_dur=0.8", "-t", f"{dur:.2f}", "-r", "25",
                        "-c:v", "libx264", "-tune", "stillimage", "-pix_fmt", "yuv420p", "-preset", "medium", "-crf", "24",
                        "-c:a", "aac", "-b:a", "96k", "-ar", "44100", seg], check=True)
        segs.append(seg)
        print(f"  слайд {i}: {title} — {dur:.1f} с")
    lst = os.path.join(tmp, "list.txt")
    with io.open(lst, "w", encoding="utf-8") as f:
        for s_ in segs:
            f.write(f"file '{s_}'\n")
    subprocess.run([ff, "-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", lst, "-c", "copy", out], check=True)
    shutil.rmtree(tmp, ignore_errors=True)
    print("видео:", out, f"{os.path.getsize(out) / 1e6:.1f} МБ")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    a = sys.argv[1:]
    run(a[0] if a else DEF_MODEL, a[1] if len(a) > 1 else DEF_OUT)
