Confirm Nepal — Desktop
========================

An offline desktop app that generates VAT confirmation letters from a bill
list. No internet, no server, no account needed at runtime.

Run from source (for development)
----------------------------------
1. Install Python 3.10+ from python.org (the standard installer includes
   Tkinter -- if you installed Python from the Microsoft Store, Tkinter may
   be missing; the python.org installer is the safe choice).
2. In this folder:

       pip install -r requirements.txt
       python app.py

Build a standalone .exe (so end users need no Python at all)
--------------------------------------------------------------
Once `pip install -r requirements.txt` has succeeded (this includes
PyInstaller), run:

       pyinstaller --noconfirm --windowed --onefile ^
           --add-data "static;static" ^
           --name "ConfirmNepal" ^
           app.py

(On macOS/Linux, replace the `^` line continuations with `\` and the
`static;static` separator `;` with `:`.)

The finished .exe will be in `dist\ConfirmNepal.exe`. That single file is
everything an end user needs -- copy it anywhere and double-click to run.
No Python, no pip, no internet required on their machine.

Notes
-----
- `generator.py` is shared, framework-agnostic logic -- it's the same file
  used by the Confirm Nepal *web* version, so fixes made in one place are
  easy to port to the other.
- The bundled Excel template lives at `static/confirm_nepal_template.xlsx`
  (regenerate it with `python build_template.py` if you change its columns).
- Everything runs locally: the bill list, business details, and generated
  letters never leave the user's machine.
