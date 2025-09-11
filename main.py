import os
import logging
import json
import copy
import pandas as pd
from typing import List, Optional, Tuple
from dataclasses import asdict

from backend import OrderItem, perform_order_item, ui_print, lookup_gtin

# try import global browser_not_found for final report
try:
    from backend import browser_not_found  # type: ignore
except Exception:
    browser_not_found = []

# === options (unchanged) ===
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


def choose_option(options: List, prompt: str):
    print(f"\n{prompt}:")
    for i, option in enumerate(options, start=1):
        print(f"{i}. {option}")
    while True:
        choice = input("Введите номер: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(options):
            return options[int(choice)-1]
        print("Неверный выбор. Попробуйте снова.")


def print_collected(collected: List[OrderItem]):
    if not collected:
        print("\n--- Накопленные позиции: пусто ---\n")
        return
    print("\n--- Накопленные позиции ---")
    for idx, it in enumerate(collected, start=1):
        print(f"{idx}. {it.simpl_name} | {it.size} | {it.units_per_pack} уп. | GTIN {it.gtin} | к-во: {it.codes_count} | заявка: '{it.order_name}'")
    print("---------------------------\n")


def choose_delete_index(collected: List[OrderItem]) -> Optional[int]:
    if not collected:
        ui_print("Нет позиций для удаления.")
        return None
    print_collected(collected)
    inp = input("Введите номер позиции для удаления (пусто = последняя): ").strip()
    if inp == "":
        return len(collected) - 1
    if not inp.isdigit():
        ui_print("Неверный ввод.")
        return None
    idx = int(inp) - 1
    if idx < 0 or idx >= len(collected):
        ui_print("Индекс вне диапазона.")
        return None
    return idx


def safe_perform(it: OrderItem) -> Tuple[bool, str]:
    try:
        res = perform_order_item(asdict(it))
        if res is None:
            logging.error("perform_order_item вернула None")
            return False, "perform_order_item вернула None"
        if isinstance(res, tuple) and len(res) == 2:
            ok, msg = res
            return bool(ok), str(msg)
        logging.error(f"perform_order_item вернула некорректный результат: {res}")
        return False, f"Некорректный результат: {res}"
    except Exception as e:
        logging.exception("Ошибка при вызове perform_order_item")
        return False, f"Exception: {e}"


def main():
    NOMENCLATURE_XLSX = "data/nomenclature.xlsx"
    if not os.path.exists(NOMENCLATURE_XLSX):
        ui_print(f"ERROR: файл {NOMENCLATURE_XLSX} не найден.")
        return

    df = pd.read_excel(NOMENCLATURE_XLSX)
    df.columns = df.columns.str.strip()

    ui_print("=== Kontur Automation — ввод позиций ===")
    collected: List[OrderItem] = []

    while True:
        print("\nПоиск по GTIN?")
        print("1. Да")
        print("2. Нет")
        gtin_choice = input("Выбор (1/2): ").strip()

        if gtin_choice == "1":
            order_name = input("Заявка (текст, будет вставлен в 'Заказ кодов №'): ").strip()
            if not order_name:
                ui_print("Нужно ввести заявку.")
                continue
            gtin_input = input("Введите GTIN: ").strip()
            if not gtin_input:
                ui_print("GTIN пустой — отмена.")
                continue
            try:
                codes_count = int(input("Количество кодов (целое): ").strip())
            except:
                ui_print("Неверно введено количество кодов. Попробуй ещё раз.")
                continue

            it = OrderItem(
                order_name=order_name,
                simpl_name="по GTIN",
                size="не указано",
                units_per_pack="не указано",
                codes_count=codes_count,
                gtin=gtin_input,
                full_name=""
            )
            collected.append(it)
            ui_print(f"Добавлено по GTIN: {gtin_input} — {codes_count} кодов — заявка '{order_name}'")
            print_collected(collected)

        elif gtin_choice == "2":
            order_name = input("\nЗаявка (текст, будет вставлен в 'Заказ кодов №'): ").strip()
            if not order_name:
                ui_print("Нужно ввести заявку.")
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
                continue

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
            print_collected(collected)

        else:
            ui_print("Неверный выбор — попробуйте снова.")
            continue

        # меню действий
        while True:
            print("\nДействия:")
            print("1 - Ввести ещё позицию")
            print("2 - Удалить позицию (по индексу)")
            print("3 - Показать накопленные позиции")
            print("4 - Выполнить все накопленные позиции")
            print("0 - Выйти без выполнения")
            action = input("Выбор (0/1/2/3/4): ").strip()
            if action == "1":
                break
            elif action == "2":
                idx = choose_delete_index(collected)
                if idx is None:
                    continue
                removed = collected.pop(idx)
                ui_print(f"Удалена позиция #{idx+1}: {removed.simpl_name} — GTIN {removed.gtin}")
                print_collected(collected)
            elif action == "3":
                print_collected(collected)
            elif action == "4":
                # подтверждение + snapshot
                print_collected(collected)
                confirm = input(f"Подтвердите выполнение {len(collected)} задач(и)? (y/n): ").strip().lower()
                if confirm != "y":
                    ui_print("Выполнение отменено пользователем.")
                    continue

                # делаем жёсткую глубокую копию коллекции (snapshot)
                to_process = copy.deepcopy(collected)
                # сохраняем snapshot на диск для дебага
                try:
                    with open("last_snapshot.json", "w", encoding="utf-8") as f:
                        json.dump([asdict(x) for x in to_process], f, ensure_ascii=False, indent=2)
                    logging.info("Saved last_snapshot.json (snapshot of to_process).")
                except Exception:
                    logging.exception("Не удалось сохранить last_snapshot.json")

                if not to_process:
                    ui_print("Нет накопленных позиций — выходим.")
                    return

                ui_print(f"\nБудет выполнено {len(to_process)} задач(и) ПОСЛЕДОВАТЕЛЬНО.")
                ui_print("Запуск...")
                results = []
                for it in to_process:
                    ui_print(f"Запуск позиции: {it.simpl_name} | GTIN {it.gtin} | заявка '{it.order_name}'")
                    ok, msg = safe_perform(it)
                    results.append((ok, msg, it))
                    ui_print(f"[{'OK' if ok else 'ERR'}] {it.simpl_name} — {msg}")

                ui_print("\n=== Выполнение завершено ===")
                success = sum(1 for r in results if r[0])
                ui_print(f"Успешно: {success}, Ошибок: {len(results)-success}.")

                # подробный отчёт
                if any(not r[0] for r in results):
                    print("\nНеудачные позиции:")
                    for ok, msg, it in results:
                        if not ok:
                            print(f" - {it.simpl_name} | GTIN {it.gtin} | заявка '{it.order_name}' => {msg}")

                if browser_not_found:
                    print("\nGTIN, не найденные в справочнике (browser_not_found):")
                    for g in sorted(set(browser_not_found)):
                        print(" -", g)

                # можно очистить collected или оставить — оставлю для возможности повторного запуска
                return
            elif action == "0":
                ui_print("Выход без выполнения.")
                return
            else:
                ui_print("Неверный выбор. Попробуйте снова.")


if __name__ == "__main__":
    main()
