PowerPoint VBA Add-in (PPAM)

How this works
- Keep your VBA modules as text files (.bas, .cls, .frm) under src/.
- Run the PowerShell script in build/import.ps1 to import everything into a new in-memory presentation and save it as a .ppam add-in.

Folder layout
- src/modules: Standard modules (.bas)
- src/classes: Class modules (.cls)
- src/forms: UserForms (.frm + .frx)
- build: Build scripts
- out: Generated .ppam add-in

Add multiple files to the same project
- Export your existing code from VBE (Alt+F11 → right-click module → Export File…).
- Drop all .bas/.cls/.frm files into the folders under src/.
- If a form has a companion .frx, place it next to the .frm with the same base name.
- Run the build script to compile them into one add-in.

Build
- Open Windows PowerShell.
- From this folder, run:
  powershell -ExecutionPolicy Bypass -File .\build\import.ps1 -AddinName "AutoFocus"

Notes
- In PowerPoint, enable: File → Options → Trust Center → Trust Center Settings… → Macro Settings → check "Trust access to the VBA project object model".
- SaveAs type uses OpenXML Add-in (ppSaveAsOpenXMLAddin = 30) to create .ppam.
