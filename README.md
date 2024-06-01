# ExcelHelperToolsWithVBA

## Overview

This repository contains a collection of helpful tools and scripts for Excel, all powered by VBA (Visual Basic for Applications). For example, _ShapeDataManager.xlsm_ allows you to output Excel shapes to a table and update the table contents to the shapes. This tool is ideal for automating repetitive tasks, enhancing Excel's functionality, and analyzing data in a structured way.

## ShapeDataManager.xlsm

### Features and Specifications

- **Enable Excel VBA**: To use these tools, you must enable macros in Excel.
- **Output Shapes to Tables**: Capture information about shapes on the `Canvas` sheet and output it to `Connectors` and `Shapes` tables.
- **Reflect Table Contents to Shapes**: Changes made in the `Connectors` and `Shapes` tables can be reflected back to the shapes on the `Canvas` sheet.
  - Currently, only color, size, position, text, and selection state are reflected back to the shapes.

### Usage

1. **Drawing Shapes**:
   - Draw your preferred shapes on the `Canvas` sheet or copy and paste shapes from other worksheets.

2. **Updating Tables**:
   - Run the `UpdateTables` macro.
   - This will capture the information of shapes placed on the `Canvas` sheet and output it to the `Connectors` and `Shapes` tables.

3. **Utilizing Tables**:
   - Use the tables for various analyses and visualizations.
   - **Examples**:
     - Analyze using pivot tables.
       - An example is provided in the `Analyze` sheet.
     - Visualize using BI tools like Excel or Power BI.
     - Let Excel Copilot analyze the data.
     - Export to CSV and let AI tools analyze the data.

4. **Reflecting Changes Back to Shapes**:
   - Run the `UpdateCanvasShapes` macro to apply changes made in the `Connectors` and `Shapes` tables back to the shapes on the `Canvas` sheet.
   - Shapes not present in the table will be deleted from the `Canvas`.
   - If a shape exists in the table but not on the `Canvas`, an error will occur.
   - **Note**: Before running this macro, make sure to run `UpdateTables` and edit the table contents accordingly.

## Developer Tab Activation

To run these macros, you need to enable the `Developer` tab on the Excel ribbon. Follow these steps to enable the `Developer` tab:

1. Click on `File` in the top left corner of Excel.
2. Select `Options` from the menu.
3. In the `Excel Options` dialog box, select `Customize Ribbon` from the left-hand menu.
4. In the right-hand panel, check the box next to `Developer` under the `Main Tabs` section.
5. Click `OK` to save your settings.

Once the `Developer` tab is enabled, you can select and run the macros from the `Macros` option in the `Developer` tab.

## License

This repository is licensed under the MIT License.

Copyright (c) 2024 [Tuckn](https://x.com/Tuckn333)
