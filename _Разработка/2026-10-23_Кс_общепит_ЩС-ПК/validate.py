"""Независимая проверка линий ЩС-ПК по выгрузке из книги (val.json)."""
import sys, json, math
d = json.load(open(sys.argv[1], encoding='utf-8'))
out = open(sys.argv[2], 'w', encoding='utf-8')
# ПУЭ табл. 1.3.7: медь, ПВХ, кабель в воздухе, трёхжильный; для 4–5 жил × 0,92 (примечание к табл.)
IDOP3 = {1.5: 19, 2.5: 25, 4: 35, 6: 42, 10: 55, 16: 75, 25: 95, 35: 120, 50: 145, 70: 180, 95: 220}
RHO, RHO_HOT, RK = 0.0225, 0.0263, 0.015      # Ом·мм²/м: рабочий, нагретый при КЗ; переходные сопр., Ом
U1, U3 = 230.0, 400.0
ZSH = 0.14104320892320893                       # Zп ф–0 на шинах ЩС-ПК (из «Сеть_щита», через ГРЩ-ЖД)
ws = []
out.write('линия | P | фаз | L | сечение | Iр книга/моё | In | Iдоп | In≤Iдоп | ΔU% книга/моё | ΔUΣ | Iкз книга/моё | Iкз/In | ok\n')
for r in d['rows']:
    name = r['D']; P = r['E'] or 0; m = int(r['H'] or 1); L = float(r['M'] or 0); S = float(r['V'] or 0)
    cos = float(r['P'] or 0.95); sin = math.sqrt(1 - cos * cos); x0 = float(r['X'] or 0) / 1000; n5 = int(r['O'] or 3)
    In = float(r['U'] or 0)
    if name.startswith('Ввод'):
        I = float(r['T'])          # ввод: Iр щита по наибольшей фазе (из «Нагрузка_щитов»)
    else:
        I = P * 1000 / ((U1 if m == 1 else math.sqrt(3) * U3) * cos)
    idop = IDOP3[S] * (0.92 if n5 >= 4 else 1.0)
    if m == 1:
        du = 2 * L * I * (RHO / S * cos + x0 * sin) / U1 * 100
    else:
        du = math.sqrt(3) * L * I * (RHO / S * cos + x0 * sin) / U3 * 100
    dusum = du + (0 if name.startswith('Ввод') else 0.97)
    zp = (0 if name.startswith('Ввод') else ZSH) + 2 * L * RHO_HOT / S + (0 if name.startswith('Ввод') else RK)
    ikz = 0.95 * U1 / zp if not name.startswith('Ввод') else None
    flags = []
    if In < I: flags.append('In<Iр')
    if In > idop + 1e-6: flags.append('кабель не защищён')
    if dusum > 4: flags.append('ΔU>4%')
    if ikz and ikz < 10 * In: flags.append('Iкз<10·In (C)')
    bi = r['BI']
    ws.append((name, flags))
    out.write(f"{name} | {P} | {m} | {L:.0f} | {n5}×{S:g} | {r['T']:.2f}/{I:.2f} | {In:g} | {idop:.1f} | {'да' if In <= idop + 1e-6 else 'НЕТ'} | "
              f"{r['AB']}/{du:.2f} | {dusum:.2f} (книга {r['BG']}) | {bi:.0f}/{(ikz or 0):.0f} | {((ikz or 0) / In if In else 0):.1f} | "
              f"{'ok' if not flags else ', '.join(flags)} (книга {r['BK']})\n")
bad = [w for w in ws if w[1]]
out.write(f'\nлиний {len(ws)}, с замечаниями {len(bad)}: {bad}\n')
