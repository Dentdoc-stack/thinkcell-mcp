"""Set up and run the think-cell remote connector (Claude in PowerPoint, claude.ai).

    python connector.py setup       one-time: ngrok, fixed address, connect code, start at login
    python connector.py start       start it in the background now (no window)
    python connector.py stop        stop it
    python connector.py status      is it running and reachable?
    python connector.py code        show the connect code (--new makes a new one)
    python connector.py signout     sign every connected Claude out (they must Connect again)
    python connector.py assist      settings for think-cell Assist (--new makes a new key)
    python connector.py autostart on|off

``run`` is what ``start`` and the login shortcut call: it keeps ``server.py
--http`` and the ngrok tunnel running and restarts either if it stops. Its
log is ``connector.log`` in the config folder. See CONNECTOR.md.
"""
from __future__ import annotations

import argparse
import getpass
import io
import json
import os
import shutil
import ssl
import subprocess
import sys
import time
import urllib.request
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
CONFIG_DIR = Path(os.environ.get("THINKCELL_CONNECTOR_HOME")
                  or Path(os.environ.get("APPDATA", Path.home())) / "thinkcell-connector")
CONFIG_FILE = CONFIG_DIR / "config.json"
STATE_FILE = CONFIG_DIR / "oauth_state.json"
RUN_FILE = CONFIG_DIR / "running.json"
LOG_FILE = CONFIG_DIR / "connector.log"
NGROK_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "ngrok"
NGROK_ZIP = "https://bin.ngrok.com/c/bNyj1mQVY4c/ngrok-v3-stable-windows-amd64.zip"
STARTUP_LINK = (Path(os.environ.get("APPDATA", Path.home())) / "Microsoft" / "Windows" / "Start Menu"
                / "Programs" / "Startup" / "think-cell connector.lnk")
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
DEFAULT_PORT = 8765


def load_config() -> dict:
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8")) if CONFIG_FILE.is_file() else {}
    except ValueError:  # unreadable: same as not set up, the commands then say to run setup
        return {}


def save_config(config: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    tmp = CONFIG_FILE.with_suffix(".tmp")  # replaced in one step: never a half-written file
    tmp.write_text(json.dumps(config, indent=2), encoding="utf-8")
    from oauth import replace_file

    replace_file(tmp, CONFIG_FILE)


def _pythonw() -> str:
    """pythonw.exe next to this Python: runs without a console window."""
    exe = Path(sys.executable)
    windowless = exe.with_name("pythonw.exe")
    return str(windowless if windowless.is_file() else exe)


def _ngrok(config: dict) -> str | None:
    for candidate in (config.get("ngrok"), shutil.which("ngrok"), NGROK_DIR / "ngrok.exe"):
        if candidate and Path(candidate).is_file():
            return str(candidate)
    return None


def _get(url: str) -> int | None:
    """HTTP status of a GET, or None when unreachable (Windows certificate store)."""
    try:
        with urllib.request.urlopen(url, timeout=10, context=ssl.create_default_context()) as reply:
            return reply.status
    except urllib.error.HTTPError as exc:
        return exc.code
    except OSError:
        return None


def _ask(prompt: str, default: str = "") -> str:
    answer = input(f"{prompt}{f' [{default}]' if default else ''}: ").strip()
    return answer or default


# ---------------------------------------------------------------- setup
def setup() -> None:
    from converter import find_ppttc_exe, find_thinkcell_dir
    from oauth import new_connect_code

    config = load_config()
    print("think-cell connector setup\n")

    ppttc = Path(find_ppttc_exe(find_thinkcell_dir()))
    print(f"think-cell: {'found, ' + str(ppttc.parent) if ppttc.is_file() else 'NOT FOUND - install think-cell first'}")

    ngrok = _ngrok(config)
    if not ngrok:
        print("\nThe connector needs ngrok (free) for a fixed https address.")
        manual = ("Install ngrok from the Microsoft Store (search 'ngrok') or https://ngrok.com/download,\n"
                  "then run setup again. (Antivirus tools often block ngrok downloads; allow it if asked.)")
        if not _ask(f"Download ngrok into {NGROK_DIR}? (y/n)", "y").lower().startswith("y"):
            sys.exit(manual)
        try:
            with urllib.request.urlopen(NGROK_ZIP, timeout=180, context=ssl.create_default_context()) as reply:
                zipfile.ZipFile(io.BytesIO(reply.read())).extract("ngrok.exe", NGROK_DIR)
        except (OSError, zipfile.BadZipFile) as exc:
            sys.exit(f"The download did not work ({exc}).\n{manual}")
        ngrok = str(NGROK_DIR / "ngrok.exe")
    config["ngrok"] = ngrok
    print(f"ngrok: {ngrok}")

    print("\nSign up free at https://dashboard.ngrok.com/signup, then copy your authtoken from\n"
          "https://dashboard.ngrok.com/get-started/your-authtoken\n"
          "If pasting here does nothing (some terminals block paste into hidden input), press Enter\n"
          "and run  ngrok config add-authtoken <token>  in another terminal instead.")
    token = getpass.getpass("Paste your ngrok authtoken (input hidden; Enter keeps the saved one): ").strip()
    if token:
        subprocess.run([ngrok, "config", "add-authtoken", token], check=True, creationflags=NO_WINDOW)

    print("\nYour free fixed address is under https://dashboard.ngrok.com/domains (e.g. name.ngrok-free.app).")
    domain = _ask("ngrok domain", config.get("public_url", "").removeprefix("https://"))
    domain = domain.removeprefix("https://").removeprefix("http://").strip("/")
    if not domain:
        sys.exit("A domain is needed: claim the free one in the ngrok dashboard, then run setup again.")
    config["public_url"] = f"https://{domain}"

    print("\nFolders Claude may read data from and save to, besides this project's output folder\n"
          "(separate with ;  leave empty for none).")
    folders = _ask("Folders", ";".join(config.get("folders", [])))
    config["folders"] = [str(Path(f.strip()).expanduser()) for f in folders.split(";") if f.strip()]
    config.setdefault("port", DEFAULT_PORT)
    config.setdefault("connect_code", new_connect_code())
    config.setdefault("api_key", new_api_key())
    save_config(config)

    if _ask("\nStart the connector automatically when you log in to Windows? (y/n)", "y").lower().startswith("y"):
        autostart(True)
    stop(quiet=True)
    start()
    print(f"\nDone. Add it in Claude: Settings > Connectors > Add custom connector\n"
          f"  Name:           think-cell\n"
          f"  URL:            {config['public_url']}/mcp\n"
          f"Then click Connect and type the connect code:  {config['connect_code']}\n"
          f"(see it again any time with: python connector.py code)\n\n"
          f"For think-cell Assist (PowerPoint/Excel), run:  python connector.py assist")


# ---------------------------------------------------------------- run / start / stop
def run() -> None:
    """Keep the server and the tunnel running (what start and the login shortcut call)."""
    config = load_config()
    ngrok = _ngrok(config)
    if not config.get("public_url") or not ngrok:
        sys.exit("The connector is not set up yet: run  python connector.py setup")
    port = config.get("port", DEFAULT_PORT)
    if _get(f"http://127.0.0.1:{port}/.well-known/oauth-authorization-server") == 200:
        sys.exit("The connector is already running.")
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if LOG_FILE.is_file() and LOG_FILE.stat().st_size > 5_000_000:
        LOG_FILE.unlink()
    log = open(LOG_FILE, "a", buffering=1, encoding="utf-8")
    commands = {
        "server": [sys.executable, str(PROJECT_ROOT / "server.py"), "--http", "--port", str(port)],
        "tunnel": [ngrok, "http", str(port), "--url", config["public_url"], "--log", "stdout"],
    }
    procs: dict[str, subprocess.Popen] = {}
    started: dict[str, float] = {}
    RUN_FILE.write_text(json.dumps({"pid": os.getpid()}), encoding="utf-8")
    try:
        while True:
            for name, command in commands.items():
                proc = procs.get(name)
                if proc is not None and proc.poll() is None:
                    continue
                if proc is not None:
                    log.write(f"{time.ctime()} {name} stopped (exit {proc.returncode}); restarting\n")
                    if time.time() - started[name] < 30:  # crashing at once: don't spin
                        time.sleep(30)
                procs[name] = subprocess.Popen(command, cwd=PROJECT_ROOT, stdout=log, stderr=subprocess.STDOUT,
                                               stdin=subprocess.DEVNULL, creationflags=NO_WINDOW)
                started[name] = time.time()
            time.sleep(5)
    finally:
        for proc in procs.values():
            proc.terminate()
        RUN_FILE.unlink(missing_ok=True)


def start() -> None:
    subprocess.Popen([_pythonw(), str(PROJECT_ROOT / "connector.py"), "run"], cwd=PROJECT_ROOT,
                     creationflags=NO_WINDOW | getattr(subprocess, "DETACHED_PROCESS", 0))
    print("Starting the connector ...")
    port = load_config().get("port", DEFAULT_PORT)
    for _ in range(30):
        if _get(f"http://127.0.0.1:{port}/.well-known/oauth-authorization-server") == 200:
            print("The connector is running.")
            return
        time.sleep(1)
    print(f"It did not come up within 30 seconds; see {LOG_FILE}")


def stop(quiet: bool = False) -> None:
    if not RUN_FILE.is_file():
        if not quiet:
            print("The connector is not running.")
        return
    pid = json.loads(RUN_FILE.read_text(encoding="utf-8"))["pid"]
    subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True, creationflags=NO_WINDOW)
    RUN_FILE.unlink(missing_ok=True)
    if not quiet:
        print("The connector is stopped.")


def status() -> None:
    config = load_config()
    if not config:
        sys.exit("Not set up yet: run  python connector.py setup")
    port, public = config.get("port", DEFAULT_PORT), config["public_url"]
    local = _get(f"http://127.0.0.1:{port}/.well-known/oauth-authorization-server") == 200
    remote = _get(f"{public}/.well-known/oauth-authorization-server") == 200
    print(f"Server on this PC:    {'running' if local else 'NOT running (python connector.py start)'}\n"
          f"Public address:       {public}/mcp  ({'reachable' if remote else 'NOT reachable'})\n"
          f"Start at login:       {'on' if STARTUP_LINK.is_file() else 'off'}\n"
          f"Folders:              {'; '.join(config.get('folders', [])) or '(only the output folder)'}\n"
          f"Log:                  {LOG_FILE}")


def code(new: bool) -> None:
    from oauth import new_connect_code

    config = load_config()
    if not config:
        sys.exit("Not set up yet: run  python connector.py setup")
    if new:
        config["connect_code"] = new_connect_code()
        save_config(config)
        print("New connect code (restart the connector to use it: stop, then start):")
    print(config["connect_code"])


def new_api_key() -> str:
    import secrets

    return secrets.token_urlsafe(32)


def assist_settings(config: dict) -> str:
    """What to enter in think-cell Assist's MCP server settings (Windows and macOS)."""
    headers = {"Authorization": f"Bearer {config['api_key']}"}
    if ".ngrok" in config["public_url"]:  # free ngrok shows browsers a warning page otherwise
        headers["ngrok-skip-browser-warning"] = "1"
    headers_json = json.dumps(headers)
    url = f"{config['public_url']}/mcp"
    return (f"think-cell Assist settings (keep the key secret):\n"
            f"  Name:     think-cell charts\n  URL:      {url}\n  Headers:  {headers_json}\n\n"
            f"Windows: Group Policy 'thinkcell Assist: MCP Servers' > Enabled, then add the three values.\n"
            f"macOS:   sudo defaults write /Library/Preferences/com.microsoft.office.plist "
            f"com.think-cell.settings.assist -dict-add MCP1_Name -string 'think-cell charts' "
            f"MCP1_URL -string '{url}' MCP1_Headers -string '{headers_json}'\n"
            f"Then in PowerPoint: Insert > thinkcell > Assist > Recent History: the server shows green.")


def assist(new: bool) -> None:
    config = load_config()
    if not config:
        sys.exit("Not set up yet: run  python connector.py setup")
    if new or not config.get("api_key"):
        config["api_key"] = new_api_key()
        save_config(config)
        print("New Assist key (restart the connector to use it: stop, then start).\n")
    print(assist_settings(config))


def signout() -> None:
    STATE_FILE.unlink(missing_ok=True)
    print("Signed out every connection. Restart the connector; Claude must Connect again.")


def autostart(on: bool) -> None:
    """A shortcut in the Windows Startup folder: no admin rights or scheduled task needed."""
    if not on:
        STARTUP_LINK.unlink(missing_ok=True)
        print("Start at login: off")
        return
    STARTUP_LINK.parent.mkdir(parents=True, exist_ok=True)
    script = ("$s = (New-Object -ComObject WScript.Shell).CreateShortcut($env:LINK); "
              "$s.TargetPath = $env:TARGET; $s.Arguments = $env:ARGS; "
              "$s.WorkingDirectory = $env:WORKDIR; $s.Save()")
    env = {**os.environ, "LINK": str(STARTUP_LINK), "TARGET": _pythonw(),
           "ARGS": f'"{PROJECT_ROOT / "connector.py"}" run', "WORKDIR": str(PROJECT_ROOT)}
    subprocess.run(["powershell", "-NoProfile", "-Command", script], env=env, check=True, creationflags=NO_WINDOW)
    print("Start at login: on")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    for name in ("setup", "run", "start", "stop", "status", "signout"):
        sub.add_parser(name)
    sub.add_parser("code").add_argument("--new", action="store_true", help="make a new connect code")
    sub.add_parser("assist").add_argument("--new", action="store_true", help="make a new Assist key")
    sub.add_parser("autostart").add_argument("state", choices=["on", "off"])
    args = parser.parse_args()
    if args.command == "code":
        code(args.new)
    elif args.command == "assist":
        assist(args.new)
    elif args.command == "autostart":
        autostart(args.state == "on")
    else:
        {"setup": setup, "run": run, "start": start, "stop": stop, "status": status, "signout": signout}[args.command]()


if __name__ == "__main__":
    main()
