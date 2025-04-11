# Altium Excel to Query Builder

A desktop application for converting between Excel pivot tables and Altium Designer rule files (.RUL). This tool simplifies the creation and management of PCB design rules by allowing engineers to work with familiar Excel spreadsheets.

## Features

- **Excel Integration**: Import/export Excel files with pivot tables containing net class clearance rules
- **Rule Management**: Create, edit, and manage Altium design rules through an intuitive interface
- **Multiple Rule Types**: Support for various rule types:
  - Electrical Clearance
  - Short-Circuit
  - Un-Routed Net
  - Un-Connected Pin
  - Modified Polygon
  - Creepage Distance
- **Bidirectional Conversion**: Generate .RUL files from Excel data or populate Excel spreadsheets from existing .RUL files
- **Modern Interface**: Clean, themeable PyQt5-based GUI with light and dark modes
- **Unit Conversion**: Automatic conversion between mil, mm, and inch units

## Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### From Source

1. Clone the repository:
```bash
git clone https://github.com/your-username/altium-excel-to-query-builder.git
cd altium-excel-to-query-builder
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python src/main.py
```

### Using pip (coming soon)

```bash
pip install altium-rule-generator
altium-rule-generator
```

## Usage

### Importing from Excel

1. Create an Excel file with your clearance rules in a pivot table format
   - Rows and columns should be net class names
   - Cell values should be clearance values in your preferred unit (mil, mm, inch)

2. In the application, go to File > Import > Import Excel File and select your Excel file
   - If the file has multiple sheets, you'll be prompted to select the sheet to import
   - The application will attempt to detect the unit type automatically

3. Review and edit the imported data in the pivot table view
   - You can change units using the dropdown if needed
   - Modify cell values directly in the table

4. Generate Altium rules using the "Generate Rules" button
   - Rules will be created based on the pivot table data
   - You can review and edit individual rules in the rule editor

5. Export to .RUL file using File > Export > Export RUL File

### Importing from RUL

1. In the application, go to File > Import > Import RUL File and select your .RUL file

2. Review imported rules in the rule editor
   - Rules will be organized by type in separate tabs
   - You can edit rule properties in the rule editor

3. Convert to pivot table view using the "Generate Pivot" button
   - The application will create a pivot table based on the clearance rules

4. Export to Excel using File > Export > Export to Excel

## Theme Support

The application supports both light and dark themes. Switch between themes using View > Theme in the menu.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
