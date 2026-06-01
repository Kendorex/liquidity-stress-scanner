from pathlib import Path
import subprocess
import sys

PROJECT_ROOT = Path(__file__).resolve().parent

SCRIPTS = [
    "preparation_data/m1/script_for_m1_signals.py",
    "preparation_data/m2/script_for_m2_signals.py",
    "preparation_data/m3/script_for_m3_signals.py",
    "preparation_data/m4/script_for_m4_signals.py",
    "preparation_data/m5/script_for_m5_signals_updated.py",
    "preparation_data/lsi/script_for_lsi_signals_fixed_v3.py",
]


def run_script(relative_path: str) -> None:
    script_path = PROJECT_ROOT / relative_path
    print("=" * 90)
    print(f"Запускаю: {relative_path}")
    print("=" * 90)

    if not script_path.exists():
        raise FileNotFoundError(f"Скрипт не найден: {script_path}")

    result = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=PROJECT_ROOT,
        text=True,
    )

    if result.returncode != 0:
        raise RuntimeError(f"Скрипт завершился с ошибкой: {relative_path}")


def main() -> None:
    print("Liquidity Stress Scanner — единый пересчёт M1–M5 и LSI")
    for script in SCRIPTS:
        run_script(script)
    print("=" * 90)
    print("Готово: все модули пересчитаны.")
    print("Итоговый файл LSI: data/lsi/results/lsi_signals.xlsx")
    print("=" * 90)


if __name__ == "__main__":
    main()
