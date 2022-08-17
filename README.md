# ODK-Build-RPA

These are RPA (robotic process automation) routines for extract data from SAP, treat all the data, save as xls file, and than convert in xml ODK Build form, in three simple steps.

## App log and run macro:
app_log_n_run_macro.py is a python script used for signing in SAP and run a excel macro saved in module 001.

## Module 001:
Module001.bas is a excel VBA script used for:
 - access SAP;
 - export IE03 report (IE03 is used to display a list of equipments);
 - export MB52 report (MB52 is used to display warehouse stocks of material);
 - treat all the data;
 - save files as a xlsform;
 - and run a python script to convert xls in xml (app_convert_xml.py).

## App convert xml:
app_convert_xml.py is a python script used for accessing web browser through selenium and converting xlsforms in ODK build xml forms.
