# VBA Stock Analysis Assignment ( Irfan Sencer Senyurt )

## Overview

This VBA script is designed to analyze stock data for specific years in Excel. Both versions of the script calculate and display the "Greatest % Increase", "Greatest % Decrease", and "Greatest Total Volume" values for each sheet. The first version also summarizes the "Yearly Change", "Percent Change", and "Total Volume" values for individual year sheets (2018, 2019, 2020).

## Usage

1. Open your Excel file and access the Visual Basic for Applications (VBA) editor by pressing `ALT + F11`.
2. Click `Insert` in the toolbar, then select `Module` to create a new module.
3. Choose one of the provided VBA scripts and copy it.
4. Paste the chosen script into the newly created module.
5. Press `CTRL + S` to save the module.
6. Close the VBA editor.
7. Press `ALT + F8` in Excel to open the "Macro" dialog box.
8. Select the `irfan_senyurt_assignment` macro and click "Run".
9. The script will process the data on each sheet (2018, 2019, 2020) and display the "Greatest % Increase", "Greatest % Decrease", and "Greatest Total Volume" values in cells Q2, Q3, and Q4, respectively, for each sheet. The first version of the script will also populate the "Yearly Change", "Percent Change", and "Total Volume" values in columns I, J, and L, respectively.

## Customization

If you need to analyze stock data for different years or modify the script's behavior, you can edit the VBA script directly in the VBA editor. For instance, you can change the array of years in the `For Each sheetname In Array("2018", "2019", "2020")` line to include different years or add more years.

## Troubleshooting

If you encounter any issues or errors while running the script, please check the following:

- Ensure the Excel file contains the required sheets with the correct naming format (e.g., 2018, 2019, 2020).
- Verify that the stock data in each sheet is organized correctly, with headers in the first row and data starting in the second row.
- Check for any syntax errors or typos in the VBA script.
