# Excel Mimic Console Application

## Overview

The Excel Mimic Console is a C# console application designed to replicate essential features of Excel. This application supports a variety of functionalities ranging from basic to complex cross-sheet formulas and includes capabilities for CSV data import and export. Powered by WebSocket technology, users can engage in real-time multi-user collaborations, making data manipulation a seamless experience.

## Features

- **Formulas:** Supports basic to complex cross-sheet formulas, enabling dynamic data manipulation.
- **CSV Import/Export:** Allows users to import data from and export data to CSV files, facilitating easy data migration.
- **Real-Time Collaboration:** WebSocket technology enables multiple users to work on the same sheet in real time.
- **Pivot Tables:** Users can create simplified pivot tables with SUM, AVG, and COUNT operations, aiding in data analysis and representation.
- **Role Switching:** Seamlessly switch roles between server and client to enhance the collaborative experience.

## Usage

### Prerequisites

- .NET Core 3.1 or later
- A compatible IDE, such as Visual Studio or Visual Studio Code

### Running the Application

1. Clone the repository to your local machine.
   ```sh
   git clone https://github.com/CristiCM/workbook_project.git

2. Navigate to the project directory.
   ```sh
   cd workbook_project

3. Run the application.
   ```sh
   dotnet run

## Keyboard Shortcuts and Formulas

Press `CTRL + L` to view the keys legend and available formulas within the application. Here is a quick reference:

### Keys Legend:

- New Sheet - `CTRL + N`
- Save Sheet - `CTRL + S`
- Open Sheet - `CTRL + O`
- Cut - `CTRL + X`
- Copy - `CTRL + C`
- Paste - `CTRL + V`
- Edit Existing Cell - `F2`
- New Sheet - `F5`
- Previous Sheet - `F6`
- Next Sheet - `F7`
- Delete Sheet - `F8`

### Formulas:

- Cell Reference `=A4`
- Sum `=SUM(A1,B1,C1)`
- Average `=AVERAGE(A1,B1,C1)`
- Count `=COUNT(A1,B1,C1)`
- Mod `=MOD(number, divisor)`
- Power `=POWER(number, power)`
- Ceiling `=CEILING(number, significance)`
- Floor `=FLOOR(number, significance)`
- Concat `=CONCATENATE(A1, B1, C1)`
- Length `=LEN(A1)`
- Replace `=REPLACE(old_text, start_index, num_chars, new_text)`
- Substitute `=SUBSTITUTE(text, old_text, new_text, instance_num (optional))`
- Now `=NOW()`
- Today `=TODAY()`
- Vlookup `=VLOOKUP(lookup_value, table_array, col_index_num)`
- Subtotal `=SUBTOTAL(func_index, A1,B1,C1)`

#### Subtotal Function Index

| Index | Function |
|-------|----------|
| 1     | AVG      |
| 2     | COUNT    |
| 9     | SUM      |

Press "Enter" to return to the sheet.

### Pivot Table Interaction:

Press `CTRL + P` to open the pivot table menu. Navigate through the menu using arrow keys and make selections or edits with the space bar.

The menu is structured as follows:

Navigate with arrow keys and press space to select fields or edit them:

Range... `A1:D5`

Row Field

[`x`] Product
[ ] Turnover
[ ] Profit
[ ] ROI

Value Fields

[ ] Product
[`x`] Turnover [x] SUM [ ] AVG [ ] COUNT
[`x`] Profit [ ] SUM [x] AVG [ ] COUNT
[`x`] ROI [ ] SUM [ ] AVG [x] COUNT

Location... `H1`

Exit...

- **Range**: Specifies the cell range for the pivot table.
- **Row Field**: Select the row fields by navigating with arrow keys and selecting with space bar.
- **Value Fields**: Navigate and select the value fields and the desired calculations (SUM, AVG, COUNT).
- **Location**: Choose where you want the pivot table to be placed.

Press "Enter" to confirm selections and exit the menu.
