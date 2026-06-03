#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка расчёта raschet_B2B.md и расчёт точки безубыточности (сколько
внедрений/подключений окупает бюджет стратегии). Все суммы — млн руб.
Запуск: python3 _verify_b2b_breakeven.py
"""

# ---------- 1. Бюджет стратегии (Таблица А raschet_B2B) ----------
BUDGET_PHASES = {
    "ФОТ команды (маркетолог 2,0 + аналитик 2,0 + дизайнер 1,5)": [2.75, 5.50, 8.25],
    "Переоснащение face2.ru (CTA, ROI-калькулятор)":              [1.50, 0.0,  0.0],
    "Брендинг и SMM (зонтик «Мигом»)":                            [1.00, 0.0,  0.0],
    "Контент-студия и кейсы":                                     [0.50, 0.0,  0.0],
    "Перформанс-реклама и SMM":                                   [1.00, 3.00, 4.00],
    "Выставки и конференции (0,3 млн/мероприятие)":               [0.60, 1.20, 1.20],
    "Независимая сертификация (NIST FRTE, iBeta L2)":             [1.00, 1.00, 0.0],
    "Референсная программа (NPS, кейсы)":                         [0.0,  0.0,  1.50],
}
phase_tot = [round(sum(s[i] for s in BUDGET_PHASES.values()), 2) for i in range(3)]
budget_total = round(sum(phase_tot), 2)
cum_budget = [round(sum(phase_tot[:i + 1]), 2) for i in range(3)]

# ---------- 2. Тарифы ГИС ЕБС (АО «ЦБТ», 29.01.2026) ----------
EBS_SKUD_PER_EMP_YEAR = 25      # руб./сотрудник/год (безлимит проходов)
PRICE_FIRST_YEAR = 1000         # руб./сотрудник, разовая плата 1-й год
PRICE_SUBSCR_YEAR = 300         # руб./сотрудник/год, со 2-го года

# ---------- 3. Unit-экономика типового объекта (5000 сотрудников) ----------
def unit_object(employees):
    rev_first = employees * PRICE_FIRST_YEAR / 1e6
    rev_subscr = employees * PRICE_SUBSCR_YEAR / 1e6
    cost_ebs = employees * EBS_SKUD_PER_EMP_YEAR / 1e6
    return {
        "rev_first": rev_first,
        "rev_subscr": rev_subscr,
        "cost_ebs": cost_ebs,
        "gp_first": rev_first - cost_ebs,
        "gp_subscr": rev_subscr - cost_ebs,
        "margin_subscr": (rev_subscr - cost_ebs) / rev_subscr * 100,
    }

# ---------- 4. Сценарная выручка (Таблица В raschet_B2B) ----------
# Уровень 2А: число объектов СКУД нарастающим итогом по фазам и средний размер
SCEN = {
    "Π (пессим.)": {"skud_obj": [1, 3, 7],  "skud_emp": 1000,
                    "tx_obj": [5, 15, 40],   "tx_rev": 0.1,
                    "kbs": [0, 0, 0]},
    "Р (реалист.)": {"skud_obj": [2, 10, 30], "skud_emp": 2000,
                     "tx_obj": [10, 90, 250], "tx_rev": 0.2,
                     "kbs": [0, 0, 1]},          # 1 сделка в фазе 3
    "О (оптим.)": {"skud_obj": [6, 26, 76], "skud_emp": 3000,
                   "tx_obj": [50, 250, 800], "tx_rev": 0.3,
                   "kbs": [0, 1, 3]},            # 1 в ф.2, 3 накопл. в ф.3
}
KBS_DEAL = 30.0  # млн руб. за сделку уровня 1

def scenario_revenue(p):
    skud_price = p["skud_emp"] * PRICE_FIRST_YEAR / 1e6   # 1-й год объекта
    rev_skud = [round(n * skud_price, 2) for n in p["skud_obj"]]
    rev_tx = [round(n * p["tx_rev"], 2) for n in p["tx_obj"]]
    rev_kbs = [round(n * KBS_DEAL, 2) for n in p["kbs"]]
    total = [round(rev_skud[i] + rev_tx[i] + rev_kbs[i], 2) for i in range(3)]
    return rev_kbs, rev_skud, rev_tx, total

# ---------- 5. Печать ----------
print("=" * 70)
print("БЮДЖЕТ СТРАТЕГИИ (Таблица А)")
print("=" * 70)
for name, s in BUDGET_PHASES.items():
    print(f"  {name:<52}{s[0]:>5}{s[1]:>6}{s[2]:>6}")
print(f"  {'Итого за фазу':<52}{phase_tot[0]:>5}{phase_tot[1]:>6}{phase_tot[2]:>6}")
print(f"  {'Кумулятивный бюджет':<52}{cum_budget[0]:>5}{cum_budget[1]:>6}{cum_budget[2]:>6}")
print(f"  ИТОГО БЮДЖЕТ 36 мес.: {budget_total} млн руб.")

print("\n" + "=" * 70)
print("UNIT-ЭКОНОМИКА ОБЪЕКТА (Таблица Б)")
print("=" * 70)
for emp in (1000, 2000, 3000, 5000):
    u = unit_object(emp)
    print(f"  {emp} сотр.: выручка 1-й год {u['rev_first']:.3f}; абонплата/год "
          f"{u['rev_subscr']:.3f}; ЕБС {u['cost_ebs']:.3f}; "
          f"ВП 1-й год {u['gp_first']:.3f}; маржа абонплаты {u['margin_subscr']:.1f}%")

print("\n" + "=" * 70)
print("СЦЕНАРНАЯ ВЫРУЧКА (кумулятивно по фазам, Таблица В), млн руб.")
print("=" * 70)
scen_total = {}
for name, p in SCEN.items():
    rev_kbs, rev_skud, rev_tx, total = scenario_revenue(p)
    scen_total[name] = total
    print(f"\n  {name}  (объект СКУД {p['skud_emp']} сотр. → "
          f"{p['skud_emp']*PRICE_FIRST_YEAR/1e6:.1f} млн/объект)")
    print(f"    Ур.1 КБС     : {rev_kbs}")
    print(f"    Ур.2А СКУД   : {rev_skud}")
    print(f"    Ур.2Б транз. : {rev_tx}")
    print(f"    ИТОГО        : {total}")

print("\n" + "=" * 70)
print("ROI И ОКУПАЕМОСТЬ (Таблица Г)")
print("=" * 70)
print(f"  {'':<14}{'Фаза1(м6)':>12}{'Фаза2(м18)':>12}{'Фаза3(м36)':>12}")
print(f"  {'Кум. бюджет':<14}{cum_budget[0]:>12}{cum_budget[1]:>12}{cum_budget[2]:>12}")
for name, total in scen_total.items():
    net = [round(total[i] - cum_budget[i], 2) for i in range(3)]
    roi = [round(net[i] / cum_budget[i] * 100) for i in range(3)]
    print(f"  Чист.эфф {name[:1]:<5}{net[0]:>12}{net[1]:>12}{net[2]:>12}")
    print(f"  ROI %    {name[:1]:<5}{roi[0]:>12}{roi[1]:>12}{roi[2]:>12}")

print("\n" + "=" * 70)
print("ТОЧКА БЕЗУБЫТОЧНОСТИ: сколько объектов окупает бюджет")
print("=" * 70)
print(f"  Полный бюджет = {budget_total} млн руб.; бюджет фазы 1 = {cum_budget[0]} млн руб.\n")
for emp in (1000, 2000, 3000, 5000):
    rev = emp * PRICE_FIRST_YEAR / 1e6
    n_full = budget_total / rev
    n_ph1 = cum_budget[0] / rev
    print(f"  Объект {emp} сотр. ({rev:.1f} млн/объект): "
          f"полный бюджет → {n_full:.1f} объектов; фаза 1 → {n_ph1:.1f} объектов")
print(f"\n  В пересчёте на сотрудников: {budget_total} млн ÷ 1000 руб. = "
      f"{budget_total*1e6/PRICE_FIRST_YEAR:,.0f} сотрудников")
print(f"  Через уровень 1: 1 сделка КБС ({KBS_DEAL} млн) покрывает "
      f"{KBS_DEAL/budget_total*100:.0f}% бюджета 36 мес.")
