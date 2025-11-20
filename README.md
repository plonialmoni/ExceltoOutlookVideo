# Excel to Outlook Appointment Creator

This project provides a VBA script to automate the creation of Outlook appointments based on data stored in an Excel spreadsheet.

## Overview

The `AddAppointmentsToOutlookCalendar` macro reads appointment details from "Sheet1" of the Excel workbook and creates corresponding events in your default Outlook calendar.

## Files

- **execol.vba**: The VBA source code for the macro.
- **Appointments VBA - Public.xlsm**: A macro-enabled Excel workbook template ready for use.
- **Appointments VBA - Public - With Color Category.xlsm**: A variant of the template (including color categorization features).

## Prerequisites

- Microsoft Excel
- Microsoft Outlook (must be installed and configured)

## Data Format

The script expects data in **Sheet1** starting from row 2 (row 1 is for headers). The columns should be arranged as follows:

| Column | Field | Description |
|--------|-------|-------------|
| **A** | Subject | The title of the appointment |
| **B** | Start Date | Date the appointment starts (e.g., 2023-10-27) |
| **C** | Start Time | Time the appointment starts (e.g., 14:00) |
| **D** | End Date | Date the appointment ends |
| **E** | End Time | Time the appointment ends |
| **F** | Location | Location of the meeting/event |
| **G** | Body | Description or notes for the appointment |

## Usage

1. **Open the Workbook**: Open `Appointments VBA - Public.xlsm`.
2. **Enter Data**: Fill in the appointment details in "Sheet1" following the format above.
3. **Run the Macro**:
   - Go to the **Developer** tab (or press `Alt + F11` to open the VBA editor).
   - If using the VBA editor, ensure the code from `execol.vba` is present in a module.
   - Run the subroutine `AddAppointmentsToOutlookCalendar`.
4. **Confirmation**: A message box will appear indicating how many appointments were successfully added.

## Notes

- The script uses late binding (`CreateObject("Outlook.Application")`), so no manual reference to the Outlook object library is required in the VBA editor.
- Ensure Outlook is running or configured to open.
