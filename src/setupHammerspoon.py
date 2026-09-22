"""
Set up / check the Hammerspoon integration used by the New-Outlook features
(RSVP, the selected-email read behind snooze/save, and vim-style j/k/h/l
message-list navigation).

Safe parts are automated; the rest is verified and reported:
  1. Copies the bundled hammerspoon/*.lua modules (outlook_rsvp, outlook_keys)
     into ~/.hammerspoon/ (non-destructive: only ever touches those filenames,
     backs up an existing different copy to .bak; never edits your init.lua).
  2. Reloads Hammerspoon if the module was (re)installed and IPC works.
  3. Verifies: Hammerspoon installed, hs CLI + IPC responding, module loads,
     and Accessibility granted to Hammerspoon — printing exact next steps for
     anything missing.

Output: a plain-text report (shown by Alfred in Large Type).
"""

import os
import time
import filecmp
import shutil
import subprocess

HS_APP = "/Applications/Hammerspoon.app"
HS_BIN = "/Applications/Hammerspoon.app/Contents/Frameworks/hs/hs"
HS_DIR = os.path.expanduser("~/.hammerspoon")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HS_SRC_DIR = os.path.join(SCRIPT_DIR, "hammerspoon")

# Lua modules the workflow ships. outlook_rsvp: RSVP + selected-email read
# (snooze/save); outlook_keys: vim-style j/k/h/l message-list navigation.
MODULES = ["outlook_rsvp", "outlook_keys"]

INIT_SNIPPET = (
    'Add these lines to ~/.hammerspoon/init.lua, then reload Hammerspoon:\n'
    '    require("hs.ipc")                             -- enables the `hs` command-line calls\n'
    '    require("outlook_keys")                       -- vim-style j/k/h/l navigation\n'
    '    require("outlook_rsvp").bindDefaultHotkeys()  -- optional: Cmd+Shift+A/T/D/X RSVP shortcuts'
)


def _hs(lua, timeout=10):
    """Run a Lua snippet via the hs CLI. Returns (ok, stdout)."""
    try:
        r = subprocess.run(
            [HS_BIN, "-c", lua], capture_output=True, text=True, timeout=timeout
        )
        return r.returncode == 0, r.stdout.strip()
    except Exception:
        return False, ""


def main():
    lines = []

    # 1. Hammerspoon installed?
    if not os.path.isdir(HS_APP):
        lines.append("[x] Hammerspoon is NOT installed.")
        lines.append("    Install it from https://www.hammerspoon.org and re-run this.")
        print("\n".join(lines))
        return
    lines.append("[ok] Hammerspoon is installed.")

    # 2. Install / update our modules (only filenames we own; never init.lua).
    module_changed = False
    for name in MODULES:
        bundled = os.path.join(HS_SRC_DIR, name + ".lua")
        target = os.path.join(HS_DIR, name + ".lua")
        if not os.path.isfile(bundled):
            lines.append(f"[x] Bundled {name}.lua missing from the workflow — cannot install.")
            continue
        os.makedirs(HS_DIR, exist_ok=True)
        if os.path.isfile(target):
            if filecmp.cmp(bundled, target, shallow=False):
                lines.append(f"[ok] {name}.lua already up to date in ~/.hammerspoon/")
            else:
                shutil.copy2(target, target + ".bak")
                shutil.copy2(bundled, target)
                module_changed = True
                lines.append(f"[ok] Updated {name}.lua in ~/.hammerspoon/ (previous saved as .bak)")
        else:
            shutil.copy2(bundled, target)
            module_changed = True
            lines.append(f"[ok] Installed {name}.lua into ~/.hammerspoon/")

    # 3. hs CLI + IPC working?
    ipc_ok, _ = _hs("return 'ok'")
    if not ipc_ok:
        lines.append("[x] The `hs` command-line is not responding (IPC not enabled).")
        lines.append("    " + INIT_SNIPPET.replace("\n", "\n    "))
        print("\n".join(lines))
        return
    lines.append("[ok] hs command-line (IPC) is working.")

    # 4-5. Read-only checks first — run them BEFORE any reload, since hs.reload()
    # briefly invalidates the IPC port and would make follow-up calls fail.
    for name in MODULES:
        ok, out = _hs(f"local ok,err = pcall(require,'{name}'); "
                      "return ok and 'module-ok' or ('module-err: '..tostring(err))")
        if ok and out == "module-ok":
            lines.append(f"[ok] {name} module loads.")
        else:
            lines.append(f"[x] {name} module did not load: {out or 'unknown error'}")

    ok, out = _hs("return hs.accessibilityState()")
    if ok and out == "true":
        lines.append("[ok] Accessibility is granted to Hammerspoon.")
    else:
        lines.append("[x] Accessibility is NOT granted to Hammerspoon.")
        lines.append("    Grant it in System Settings > Privacy & Security > Accessibility.")

    # 6. Activate a freshly-(re)installed module LAST, then wait for Hammerspoon
    # to come back (reload restarts the Lua state and drops the IPC port).
    if module_changed:
        _hs("hs.reload()")
        back = False
        for _ in range(50):  # wait up to ~5s for the IPC port to come back
            if _hs("return 'ok'", timeout=3)[0]:
                back = True
                break
            time.sleep(0.1)
        lines.append("[ok] Reloaded Hammerspoon to activate the module."
                     if back else "[..] Reloaded Hammerspoon — it is still restarting.")

    print("\n".join(lines))


if __name__ == "__main__":
    main()
