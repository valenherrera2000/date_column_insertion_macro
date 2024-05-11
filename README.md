# Excel Date Column Insertion Macro

## Overview

This VBA macro automatically inserts a new date column in Excel sheets

## Purpose

The purpose of this macro is to streamline new date creation in daily sheets and trackers.

## Customization
1. Copy and paste this code into your VBA editor.
2. Modify the parameters passed to `AddNewDate` within the `NewDate_Macro` subroutine.
- Example: `AddNewDate "Task", 4` (new column will be added in row 4, next to "Task" cell)
3. Embed the macro in a button or keyboard shortcut.

## Subroutines

- **NewDate_Macro:** Main subroutine responsible for executing the macro.
- **Confirmation_MsgBox:** Displays a confirmation message box before proceeding with the macro execution.
- **Withdrawal_MsgBox:** Displays a message box indicating that the macro was not run.
- **Success_MsgBox:** Displays a message box with the execution time if the macro runs successfully.
- **ScreenUpdating:** Controls the screen updating feature to improve macro performance.
- **AddNewDate:** Subroutine to perform the main action of adding a new date column. Parameters: **headerToFind**, **rowToSearch**


## Additional Notes
- Expand upon this macro to accommodate your project needs.
