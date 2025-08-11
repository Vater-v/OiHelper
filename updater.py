# updater.py  -- минималистичный, надёжный апдейтер для Windows
import argparse, os, sys, time, shutil, zipfile, hashlib, tempfile, logging
from datetime import datetime
from pathlib import Path

# --- Windows ctypes (без внешних зависимостей)
import ctypes
from ctypes import wintypes

kernel32 = ctypes.windll.kernel32
user32   = ctypes.windll.user32

WM_CLOSE = 0x0010
SW_HIDE  = 0

def _open_process(pid):
    SYNCHRONIZE = 0x00100000
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    PROCESS_TERMINATE = 0x0001
    h = kernel32.OpenProcess(SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_TERMINATE, False, pid)
    return h

def _wait_for_exit(pid, timeout_sec):
    h = _open_process(pid)
    if not h:
        return True  # не нашли — считаем завершённым
    try:
        INFINITE = 0xFFFFFFFF
        if timeout_sec <= 0:
            res = kernel32.WaitForSingleObject(h, INFINITE)
            return res == 0
        else:
            res = kernel32.WaitForSingleObject(h, int(timeout_sec * 1000))
            return res == 0
    finally:
        kernel32.CloseHandle(h)

def _enum_windows_and_post_close(pid):
    EnumWindows = user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    GetWindowThreadProcessId = user32.GetWindowThreadProcessId
    PostMessageW = user32.PostMessageW

    def callback(hwnd, lparam):
        pid_out = wintypes.DWORD()
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid_out))
        if pid_out.value == pid:
            try:
                PostMessageW(hwnd, WM_CLOSE, 0, 0)
            except Exception:
                pass
        return True

    EnumWindows(EnumWindowsProc(callback), 0)

def _try_terminate(pid):
    h = _open_process(pid)
    if not h:
        return
    try:
        kernel32.TerminateProcess(h, 1)
    finally:
        kernel32.CloseHandle(h)

def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()

def extract_zip(zip_path: Path, out_dir: Path):
    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(out_dir)

def _is_subpath(path: Path, base: Path) -> bool:
    try:
        path.resolve().relative_to(base.resolve())
        return True
    except Exception:
        return False

def main():
    ap = argparse.ArgumentParser(description="OiHelper Updater")
    ap.add_argument("--pid", type=int, required=False, help="PID основной программы")
    ap.add_argument("--zip", required=True, help="Путь к архиву обновления")
    ap.add_argument("--appdir", required=True, help="Папка приложения (куда ставить)")
    ap.add_argument("--exe", default="OiHelper.exe", help="EXE для запуска после апдейта")
    ap.add_argument("--version", default="", help="Версия для записи в version.txt")
    ap.add_argument("--sha256", default="", help="Ожидаемый sha256 архива")
    ap.add_argument("--timeout", type=int, default=25, help="Сколько ждать закрытия процесса, сек")
    ap.add_argument("--logdir", default="", help="Папка логов (по умолчанию LocalAppData)")
    args = ap.parse_args()

    # Логирование
    local = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData/Local")))
    log_dir = Path(args.logdir) if args.logdir else (local / "OiHelper" / "Updater" / "logs")
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"apply_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logging.basicConfig(filename=log_path, level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
    logging.info("=== START updater ===")
    logging.info("Args: %s", vars(args))

    zip_path = Path(args.zip)
    appdir   = Path(args.appdir)
    exe_path = appdir / args.exe
    packages = local / "OiHelper" / "Packages"
    packages.mkdir(parents=True, exist_ok=True)

    if not zip_path.exists():
        logging.error("ZIP не найден: %s", zip_path)
        print("Ошибка: ZIP не найден.")
        return 2

    # Проверка sha256 (если передана)
    if args.sha256:
        calc = sha256_file(zip_path)
        if calc.lower() != args.sha256.lower():
            logging.error("SHA256 mismatch: expected=%s actual=%s", args.sha256, calc)
            print("Ошибка: контрольная сумма не совпала.")
            return 3
        logging.info("SHA256 OK")

    # Если апдейтер запущен из appdir — скопировать себя во временный каталог и перезапуститься
    try:
        self_path = Path(sys.argv[0]).resolve()
        if _is_subpath(self_path, appdir):
            tmp_dir = Path(tempfile.mkdtemp(prefix="updater_reexec_"))
            new_self = tmp_dir / self_path.name
            shutil.copy2(self_path, new_self)
            os.execv(str(new_self), [str(new_self)] + sys.argv[1:])
            return 0
    except Exception as e:
        logging.warning("Не удалось реисполнить из временной папки: %s", e)

    # Закрыть основное приложение аккуратно
    if args.pid and args.pid > 0:
        _enum_windows_and_post_close(args.pid)
        if not _wait_for_exit(args.pid, args.timeout):
            logging.warning("Процесс не завершился за %ss. Пробую TerminateProcess.", args.timeout)
            _try_terminate(args.pid)
            time.sleep(1)

    # Распаковать во временную директорию
    stage_dir = Path(tempfile.mkdtemp(prefix="oihelper_stage_"))
    try:
        extract_zip(zip_path, stage_dir)
        logging.info("Распаковано в %s", stage_dir)
    except Exception as e:
        logging.exception("Ошибка распаковки")
        print("Ошибка распаковки:", e)
        return 4

    # Атомарная замена: appdir -> backup, stage -> appdir
    backup_root = appdir.parent / "backup"
    backup_root.mkdir(parents=True, exist_ok=True)
    backup_dir = backup_root / f"{appdir.name}._old_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    try:
        if appdir.exists():
            # попытка переименования (быстро и надёжно)
            appdir.rename(backup_dir)
            logging.info("Старую папку перенёс в %s", backup_dir)
        else:
            logging.info("Папка приложения отсутствует, устанавливаем начисто.")

        # Перенос staged в appdir
        appdir.parent.mkdir(parents=True, exist_ok=True)
        stage_final = appdir
        stage_dir.rename(stage_final)
        logging.info("Новая версия перенесена в %s", stage_final)

        # Записать версию (если задана)
        if args.version:
            try:
                (stage_final / "version.txt").write_text(args.version, encoding="utf-8")
            except Exception:
                logging.warning("Не удалось записать version.txt")

        # Запуск приложения
        try:
            os.chdir(stage_final)
            proc = ctypes.windll.shell32.ShellExecuteW(0, "open", str(exe_path), None, str(stage_final), 1)
            logging.info("ShellExecute result=%s", proc)
        except Exception as e:
            logging.exception("Не удалось запустить приложение")
            print("Обновление применено, но запуск не удался:", e)
            # не откатываем, файлы уже на месте

        # Очистить бэкап (можно отложить/оставлять n последних)
        try:
            shutil.rmtree(backup_dir, ignore_errors=True)
            logging.info("Бэкап очищен")
        except Exception as e:
            logging.warning("Не удалось удалить бэкап: %s", e)

        print("Готово: обновление применено.")
        return 0

    except Exception as e:
        logging.exception("Ошибка применения обновления")
        # Rollback
        try:
            if appdir.exists():
                shutil.rmtree(appdir, ignore_errors=True)
            if backup_dir.exists():
                backup_dir.rename(appdir)
                logging.info("Rollback успешно выполнен.")
        except Exception as re:
            logging.error("Не удалось выполнить rollback: %s", re)
        print("Ошибка: обновление не применено, выполнен откат.")
        return 5

    finally:
        # Чистка временной распаковки (если не перенеслась)
        try:
            if stage_dir.exists():
                shutil.rmtree(stage_dir, ignore_errors=True)
        except Exception:
            pass

if __name__ == "__main__":
    sys.exit(main())
