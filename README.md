# Excel Date Column Insertion Macro

## Overview
This VBA macro automatically inserts a new date column in Excel sheets, ideal for daily sheets and trackers.

## Features
- **Dynamic Date Insertion:** Inserts a new column with the **current date** formatted as "dd-mmm-yy".
- **Customizable Search:** Locates the row and header specified by the user and adds a new column to its right.
- **User Interaction:** Utilizes message boxes for confirmations and notifications about the process outcome.

## How It Works
- The macro checks for a specific header name (e.g., "Task") in a designated row (e.g., row 4).
- If found, it inserts a new column right next to this header.
- It sets the new column's header to the current date, e.g., "dd-mmm-yy".

## Customization
Change the parameters in the call to `AddNewDate` within the `NewDate_Macro` subroutine.
- Example: `AddNewDate "Task", 4`

## Additional Notes
- Embedding the macro in a button allows easy execution without navigating the VBA editor.
- Messages will inform you if the macro was not run or if the specified header wasn't found.
