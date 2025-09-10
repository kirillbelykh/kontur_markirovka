import os
import logging
import pandas as pd
from typing import List
from dataclasses import asdict
from backend import OrderItem, perform_order_item, ui_print, lookup_gtin

simplified_options = [
    "стер лат 1-хлор", "стер лат", "стер лат 2-хлор", "стер нитрил",
    "хир", "хир 1-хлор", "хир с полимерным", "хир 2-хлор", "хир изопрен",
    "хир нитрил", "ультра", "гинекология", "двойная пара", "микрохирургия",
    "ортопедия", "латекс диаг гладкие", "латекс диаг", "латекс 2-хлор",
    "латекс с полимерным", "латекс удлиненный", "латекс анатомической",
    "латекс hr", "латекс 1-хлор", "нитрил диаг", "нитрил диаг hr короткий",
    "нитрил диаг hr удлиненный"
]

color_required = [
    "латекс 1-хлор", "латекс 2-хлор", "латекс HR", "латекс анатомической",
    "латекс диаг", "латекс диаг гладкие", "латекс с полимерным",
    "латекс удлиненный", "нитрил диаг", "нитрил диаг HR короткий",
    "нитрил диаг HR удлиненный", "стер лат 1-хлор", "стер лат 2-хлор"
]

venchik_required = [
    "гинекология", "микрохирургия", "ортопедия"
]

color_options = ["белый", "зеленый", "натуральный", "розовый", "синий", "фиолетовый", "черный"]
venchik_options = ["с венчиком", "без венчика"]

size_options = [
    "XS", "S", "M", "L", "XL", "5,0", "5,5", "6,0", "6,5",
    "7,0", "7,5", "8,0", "8,5", "9,0", "9,5", "10,0"
]

units_options = [1,2,3,4,5,6,7,8,9,10,20,25,30,40,50,60,70,80,90,100,110,120,125,250,500]

def choose_option(options, prompt):
    print(f"\n{prompt}:")
    for i, option in enumerate(options, start=1):
        print(f"{i}. {option}")
    while True:
        choice = input("Введите номер: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(options):
            return options[int(choice)-1]
        print("Неверный выбор. Попробуйте снова.")

def main():
    NOMENCLATURE_XLSX = "data/nomenclature.xlsx"
    if not os.path.exists(NOMENCLATURE_XLSX):
        ui_print(f"ERROR: файл {NOMENCLATURE_XLSX} не найден.")
        return

    df = pd.read_excel(NOMENCLATURE_XLSX)
    df.columns = df.columns.str.strip()

    collected: List[OrderItem] = []

    while True:
        print("\n=== Ввод новой позиции ===")
        if collected:
            print("0 - Отменить последнюю добавленную позицию")
        order_name = input("Заявка (текст, будет вставлен в 'Заказ кодов №'): ").strip()
        if order_name == "0" and collected:
            removed = collected.pop()
            ui_print(f"✅ Последняя позиция удалена: {removed.simpl_name} ({removed.size}, {removed.units_per_pack} уп.)")
            continue

        simpl = choose_option(simplified_options, "Выберите вид товара")
        color = None
        if simpl.lower() in [c.lower() for c in color_required]:
            color = choose_option(color_options, "Выберите цвет")
        venchik = None
        if simpl.lower() in [c.lower() for c in venchik_required]:
            venchik = choose_option(venchik_options, "С венчиком/без венчика?")
        size = choose_option(size_options, "Выберите размер")
        units = choose_option(units_options, "Выберите количество единиц в упаковке")

        try:
            codes_count = int(input("Количество кодов (целое): ").strip())
        except:
            ui_print("Неверно введено количество кодов. Попробуй ещё раз.")
            continue

        gtin, full_name = lookup_gtin(df, simpl, size, units, color, venchik)
        if not gtin:
            ui_print(f"GTIN не найден для ({simpl}, {size}, {units}, {color}, {venchik}) — позиция не добавлена.")
        else:
            it = OrderItem(
                order_name=order_name,
                simpl_name=simpl,
                size=size,
                units_per_pack=units,
                codes_count=codes_count,
                gtin=gtin,
                full_name=full_name or ""
            )
            collected.append(it)
            ui_print(f"Добавлено: {simpl} ({size}, {units} уп., {color or 'без цвета'}) — GTIN {gtin} — {codes_count} кодов — заявка '{order_name}'")

        print("\n1 - Ввести ещё позицию\n2 - Выполнить все накопленные позиции\n0 - Отменить последнюю добавленную позицию")
        choice = input("Выбор: ").strip()
        if choice == "0" and collected:
            removed = collected.pop()
            ui_print(f"✅ Последняя позиция удалена: {removed.simpl_name} ({removed.size}, {removed.units_per_pack} уп.)")
            continue
        elif choice == "1":
            continue
        elif choice == "2":
            break

    if not collected:
        ui_print("Нет накопленных позиций — выходим.")
        return

    ui_print(f"\nБудет выполнено {len(collected)} задач(и) ПОСЛЕДОВАТЕЛЬНО.")
    ui_print("Запуск...")

    results = []
    for it in collected:
        try:
            ok, msg = perform_order_item(asdict(it))
            results.append((ok, msg, it))
            ui_print(f"[{'OK' if ok else 'ERR'}] {it.simpl_name} — {msg}")
        except Exception as e:
            logging.exception("Ошибка при выполнении задачи")
            results.append((False, str(e), it))
            ui_print(f"[ERR] {it.simpl_name} — exception: {e}")

    ui_print("\n=== Выполнение завершено ===")
    success = sum(1 for r in results if r[0])
    ui_print(f"Успешно: {success}, Ошибок: {len(results)-success}.")

if __name__ == "__main__":
    main()
