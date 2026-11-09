COMPATIBILITY REPORT (RedM Compat Quick v1.1.0)
===============================================
(Español: ver LEEME.txt)

It lets the asset creator see whether any other resource on your server conflicts with their asset.
IT DOES NOT CHANGE ANYTHING ON THE SERVER: it only reads files and writes a .txt report. No internet connection.
Nothing to install: it uses what Windows already includes (it does not work on Linux servers).

Steps
-----
1. Unzip this folder on the server (for example on the Desktop).
2. Double-click "Generar reporte de compatibilidad.bat".
   (If Windows shows a blue warning, click "More info" and then "Run anyway".)
3. At the top, choose the report language (Español / English).
4. In the window select:
     1) the asset folder (the resource the creator gave you),
     2) the server "resources" folder,
     3) the server .cfg files: server.cfg and, if your resources are in another file (for example
        resources.cfg), select it too (hold Ctrl to select several).
        If you leave this field empty, it looks for server.cfg by itself.
5. Click "Create report". It takes a few seconds.
6. The "reports" folder opens with the file compat_<asset>_<date>_<time>.txt.
   Send that file to the asset creator.

What the report says
--------------------
- Whether there are compatibility problems.
- The list of resources that conflict with the asset (the ones the creator will ask you for).
- Details for each conflicting file: which one loads and which one does not.
- Useful warnings: resources missing from the .cfg files, resources with the same name in two folders,
  badly built hashed file names.

DISCLAIMER
----------
This tool only READS files: it does not modify, delete or move anything on the server, does not connect to the
internet and does not send any data. It is provided "as is", without warranty of any kind. The report is
informational, is based on file names and on the selected .cfg files, and may not detect every issue (for
example, conflicts inside encrypted files or scripts). Whoever runs it does so at their own risk; the author is
not liable for any damage, loss or decision made based on the report. Back up your server before changing
anything. Only use the version the asset creator gave you through their official channel.
