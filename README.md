## Recommendation

Before starting this, it is recommended that you use code external to your workbook to manage VBA code for that workbook. This solution nests the sync module in the workbook itself, which is somewhat messy. The following library supports writing your own sync code in C#: https://gitlab.com/hectorjsmith/csharp-excel-vba-sync.

## Main

The purpose of this VBA code is to allow you to set up a VBA coding project and put it under version control e.g. Git. The primary use case is for you to use this code as it is, without modifying it, as a tool for importing/exporting your own code. You can of course also contribute to this codebase- but beware: if you are editing the sync module itself, you will not be able to use it to do sync- it can't import/export itself! More on that below. The below instructions are for you as a user of this codebase, for importing/exporting your own code- not as a developer/contributor to this codebase!

## ENVIRONMENT SETUP INSTRUCTIONS

-Create your own repo and add all the files from this repo into it. You can do this e.g. by cloning this repo, deleting the remote, and then setting your own remote.

-Add a macro-enabled (.xlsm or .xlsb) workbook to the repo, in the top-level folder

-Add the All_Sync module to your excel workbook manually through the VBA editor by clicking import

-Add the following references to your workbook through the VBA editor:
* Microsoft Scripting Runtime
* Microsoft Visual Basic for Applications Extensibility

-Start coding by adding whatever modules you like. Whenever you add a module, you need to add the module name, without any extensions, to the file specificSyncList.config or genSyncList.config. These will control which of your files are exported/imported and will export/import them to/from different folders, allowing you to separate your workbook-specific files from files you want to share among multiple workbooks.

-When you want to export everything, run the export subroutine. When you want to import, run importModulesWarn (the warning is there to make you aware that you may overwrite data in the existing modules)

-Warning: do not tamper with the All_Sync module or add it to either of the config files unless you know what you're doing. It is possible to put this under version control, and export is fine, but you run into serious complications when the module tries to import itself. Also, if you tamper with this by making it dependent on other modules, import will probably break, because Excel won't be able to delete them and replace them with the newly imported files, since they are locked until the VBA macro ends. If that seems a bit confusing, the take home is: don't touch this file!
