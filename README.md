# Product Breakdown Structure (PBS)
Inspired by [MarbleMachineX](https://www.reddit.com/r/MarbleMachineX/comments/9szkud/how_to_organize_your_project_with_a_pbs_system/) and [NASA's PBS](http://web.csulb.edu/~hill/ee400d/Lectures/Week%2004%20Modeling/e_Product%20Breakdown%20Structure.pdf).

PBS document includes macros to generate a Bill Of Materials (BOM).

# Repository Files

[Example Google Spreadsheet](https://docs.google.com/spreadsheets/d/1hJnaWNOxw2grD4kduyf22FARP7wVjzjn6f_6CNK4B-Y) (license: CC BY-SA Wintergatan, Andrew Smart).
Example spreadsheet also included as '01 - Product Breakdown Structure - Dexter.ods' in this repo.

Other files in this repository are licensed by GPLv2.

macros.gs consists of javascript macros (which use the Google Apps Script API) meant for use within a Google Spreadsheet. The macros classify CAD components, assign PBS #'s to components, and generate Bill Of Materials.

export_to_pbs_csv.FCMacro consists of FreeCAD python macros meant for use within FreeCAD. It processes a FreeCAD model (possibly imported as a STEP file exported from a different CAD suite) and outputs a CSV meant for import into the PBS spreadsheet. Intended to speed up the update of the PBS document as the CAD model is updated. Tested against FreeCAD 0.17 Revision: 12595 (Git).
