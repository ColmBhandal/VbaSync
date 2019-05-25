******* ENVIRONMENT SETUP INSTRUCTIONS *******

-Clone the repo

-Add the All_Sync module to your excel workbook manually through the VBA editor by clicking import

-Add the following references to your workbook through the VBA editor:
> Microsoft Scripting Runtime
> Microsoft Visual Basic for Applications Extensibility

-Start coding by adding whatever modules you like. Whenever you add a module, you need to add the module name, without any extensions, to the file specificSyncList.config or genSyncList.config. These will export/import your files to/from different folders, allowing you to separate your workbook-specific files from files you want to share among multiple workbooks.

-Warning: do not tamper with the All_Sync module or add it to either of the config files unless you know what you're doing. It is possible to put this under version control, and export is fine, but you run into serious complications when the module tries to import itself.