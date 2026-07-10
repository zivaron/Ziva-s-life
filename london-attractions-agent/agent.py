#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
סוכן האטרקציות של זיווה בלונדון
--------------------------------
סוכן אינטראקטיבי בשורת הפקודה שממליץ על אטרקציות בלונדון,
מאפשר סינון לפי קטגוריה / אזור / תקציב, ובונה מסלול טיול
לפי מספר הימים ושעות הפנאי הפנויות.

הרצה: python3 agent.py
"""

import random

from attractions_data import ATTRACTIONS, CATEGORIES, AREAS, PRICE_LEVELS

MENU = """
=== סוכן האטרקציות של זיווה בלונדון ===
1. הצג את כל האטרקציות
2. סנן לפי קטגוריה
3. סנן לפי אזור
4. סנן לפי תקציב (חינם / בינוני / יקר)
5. בנה מסלול טיול לפי מספר ימים
6. הפתעה - אטרקציה אקראית
7. חיפוש לפי שם
0. יציאה
"""


def print_attraction(item, index=None):
    prefix = f"{index}. " if index is not None else "- "
    kids = "מתאים לילדים" if item["kids_friendly"] else "פחות מתאים לילדים"
    print(
        f"{prefix}{item['name']}\n"
        f"   קטגוריה: {item['category']} | אזור: {item['area']} | "
        f"תקציב: {item['price']} | משך: כ-{item['duration_hours']} שעות | "
        f"{kids} | דירוג: {item['rating']}\n"
        f"   {item['description']}\n"
    )


def show_all():
    for i, item in enumerate(ATTRACTIONS, 1):
        print_attraction(item, i)


def choose_from_list(options, title):
    print(f"\n{title}:")
    for i, opt in enumerate(options, 1):
        print(f"{i}. {opt}")
    choice = input("בחרי מספר: ").strip()
    if choice.isdigit() and 1 <= int(choice) <= len(options):
        return options[int(choice) - 1]
    print("בחירה לא תקינה.")
    return None


def filter_by_category():
    category = choose_from_list(CATEGORIES, "קטגוריות זמינות")
    if not category:
        return
    results = [a for a in ATTRACTIONS if a["category"] == category]
    print(f"\nנמצאו {len(results)} אטרקציות בקטגוריית '{category}':\n")
    for i, item in enumerate(results, 1):
        print_attraction(item, i)


def filter_by_area():
    area = choose_from_list(AREAS, "אזורים זמינים")
    if not area:
        return
    results = [a for a in ATTRACTIONS if a["area"] == area]
    print(f"\nנמצאו {len(results)} אטרקציות באזור '{area}':\n")
    for i, item in enumerate(results, 1):
        print_attraction(item, i)


def filter_by_price():
    price = choose_from_list(PRICE_LEVELS, "רמות תקציב")
    if not price:
        return
    results = [a for a in ATTRACTIONS if a["price"] == price]
    print(f"\nנמצאו {len(results)} אטרקציות בתקציב '{price}':\n")
    for i, item in enumerate(results, 1):
        print_attraction(item, i)


def build_itinerary():
    try:
        days = int(input("כמה ימים תרצי לתכנן? ").strip())
        hours_per_day = float(
            input("כמה שעות פנאי לאטרקציות יש בכל יום? ").strip()
        )
    except ValueError:
        print("קלט לא תקין, נסי שוב עם מספרים.")
        return

    kids_input = input("הטיול כולל ילדים? (כן/לא): ").strip()
    kids_only = kids_input in ("כן", "כ", "yes", "y")

    pool = [a for a in ATTRACTIONS if (not kids_only or a["kids_friendly"])]
    pool = sorted(pool, key=lambda a: a["rating"], reverse=True)

    total_hours_budget = days * hours_per_day
    day_plans = [[] for _ in range(days)]
    day_remaining = [hours_per_day] * days
    used_hours = 0.0

    for item in pool:
        # שיבוץ באלגוריתם "first-fit": היום הראשון שיש בו מספיק שעות פנויות
        for day_idx in range(days):
            if item["duration_hours"] <= day_remaining[day_idx]:
                day_plans[day_idx].append(item)
                day_remaining[day_idx] -= item["duration_hours"]
                used_hours += item["duration_hours"]
                break

    print(f"\n=== מסלול מוצע ל-{days} ימים ({total_hours_budget:.1f} שעות פנאי) ===\n")
    if used_hours == 0:
        print("לא נמצאו אטרקציות שמתאימות למגבלת הזמן.")
        return

    for day_idx, items in enumerate(day_plans, 1):
        print(f"--- יום {day_idx} ---")
        if not items:
            print("(אין אטרקציות מתאימות ליום זה)\n")
            continue
        for item in items:
            print_attraction(item)

    print(f"סה\"כ שעות מתוכננות: {used_hours:.1f} מתוך {total_hours_budget:.1f} שעות זמינות.")


def surprise_me():
    item = random.choice(ATTRACTIONS)
    print("\nההפתעה שלך:\n")
    print_attraction(item)


def search_by_name():
    query = input("מה לחפש? ").strip().lower()
    results = [a for a in ATTRACTIONS if query in a["name"].lower()]
    if not results:
        print("לא נמצאו תוצאות.")
        return
    for i, item in enumerate(results, 1):
        print_attraction(item, i)


def main():
    print("שלום זיווה! אני סוכן האטרקציות שלך בלונדון. איך אפשר לעזור?")
    actions = {
        "1": show_all,
        "2": filter_by_category,
        "3": filter_by_area,
        "4": filter_by_price,
        "5": build_itinerary,
        "6": surprise_me,
        "7": search_by_name,
    }
    while True:
        print(MENU)
        choice = input("בחרי אפשרות: ").strip()
        if choice == "0":
            print("טיול נעים בלונדון! להתראות.")
            break
        action = actions.get(choice)
        if action:
            action()
        else:
            print("בחירה לא תקינה, נסי שוב.")


if __name__ == "__main__":
    main()
