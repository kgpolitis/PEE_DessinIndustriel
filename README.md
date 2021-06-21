# PEE_DessinIndustriel
 Dessin Industriel
====================

A. Run "Windows Powershell" as an administrator:
    1. At "Type to search" next to your windows icon type:
      Windows Powershell
    2. Right click on the best match and choose "Run as administrator"

B. Automate the generation of pdf powerpoint files
    1. Navigate to the project's folder:

       cd c:\Your_Local_Repository

    2. Change (locally) the execution of scripts :

        Set-ExecutionPolicy RemoteSigned

    3. run the ps1 script :

        make_ppt_pdfs.ps1
