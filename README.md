***

# Ellea1 FWD Excel Automation Tool

This repository contains a Python automation script designed to batch process Falling Weight Deflectometer (FWD) data. It acts as a high-speed bridge between raw CSV input files and the **Ellea1** Excel worksheet calculator.

By using `xlwings` and `pandas`, this tool automates the tedious process of manual data entry, processing thousands of rows of layer stiffness and thickness data to calculate deflection metrics efficiently.

## Credits & Attribution

The Excel calculator software (**Ellea1**) utilized by this automation script was developed by **Associate Professor Eyal Levenberg** at the Technical University of Denmark (DTU).

*   **Author:** [Prof. Eyal Levenberg](https://orbit.dtu.dk/en/persons/eyal-levenberg/)
*   **Software License:** Freeware (Educational/Research)
*   **Citation:** Levenberg, E. (2016). *ELLEA1: Isotropic Layered Elasticity in Excel: Pavement analysis tool for students and engineers*.

> **Important:** This repository contains the *automation script only*. It does not distribute the Ellea1 Excel file. You must download the worksheet separately using the official links below.

## Features

*   **Batch Processing:** Automatically processes multiple `.csv` files found in a specified input directory.
*   **Two-Way Communication:** Writes input parameters (Thickness, Stiffness, Poisson ratio) to Excel and retrieves calculated Deflection values ($d1$ through $d10$).
*   **Data Safety:** 
    *   **Checkpointing:** Automatically saves progress every 5,000 rows to prevent data loss in case of a crash or power failure.
    *   **Memory Management:** Appends data to output files incrementally to keep RAM usage low.
*   **Headless Operation:** Runs the Excel instance in the background (`visible=False`) for faster execution.

## Prerequisites

To run this code, you need:

1.  **Microsoft Excel** installed on the local machine.
2.  **Python 3.x** installed.
3.  **Ellea1 Excel Worksheet** (Required for calculations).
    *   [Official Download (DTU Orbit)](https://orbit.dtu.dk/en/publications/ellea1-isotropic-layered-elasticity-in-excel-pavement-analysis-to)

### Python Dependencies

Install the required libraries via pip:

```bash
pip install pandas xlwings
```

## Input Data Format

For the automation to work, your input `.csv` files must use the exact headers listed below. The script maps these headers directly to specific cells in the Excel worksheet.

You can download the `sample_input.csv` file in this repository to use as a template.

### Required CSV Headers (Case Sensitive):

```text
Thickness layer 1 (mm)
Thickness layer 2 (mm)
Thickness layer 3 (mm)
Stiffness layer 1 (MPa)
Stiffness layer 2 (MPa)
Stiffness layer 3 (MPa)
Stiffness layer 4 (MPa)
Poisson ratio layer 1
```

### Example CSV Structure:

```csv
Thickness layer 1 (mm),Thickness layer 2 (mm),Thickness layer 3 (mm),Stiffness layer 1 (MPa),Stiffness layer 2 (MPa),Stiffness layer 3 (MPa),Stiffness layer 4 (MPa),Poisson ratio layer 1
150,200,300,5000,1200,400,100,0.35
160,210,310,5500,1300,450,110,0.35
```

## Configuration

Before running the script, you must update the file paths in the code to match your computer's folder structure.

Open the python script (e.g., `main.py`) and locate the `--- CONFIGURATION ---` section at the top:

```python
# 1. Path to the Ellea1 Excel file you downloaded
EXCEL_CALCULATOR_FILE = r'C:\Users\YourName\Documents\FWD\Ellea1_Worksheet.xlsx'

# 2. Directory containing your input CSVs
INPUT_FILES = glob.glob(r"C:\Users\YourName\Documents\FWD\Input\*.csv")

# 3. Directory where results will be saved
OUTPUT_DIRECTORY = r"C:\Users\YourName\Documents\FWD\Output"
```

## Usage

1.  **Prepare Data:** Place your `.csv` files in the defined input folder. Ensure headers match the requirements above.
2.  **Open Excel:** (Optional) You can leave Excel closed; the script will open it in the background.
3.  **Run Script:**

```bash
python main.py
```

**Results:** The script will generate a corresponding `_output.csv` file in your output directory, containing the original inputs plus the calculated deflection values (`d1(nm)` to `d10(nm)`).

## Troubleshooting

*   **KeyError:** If the script crashes with a `KeyError`, check your CSV headers. They must match the exact spelling and capitalization defined in the **Input Data Format** section.
*   **Excel Hanging:** If the script stops processing, ensure no Excel dialogue boxes (like "Update Links", "Activate Office", or "Recovery Pane") are open in the background.
*   **Permissions:** Ensure you have read/write access to the folders defined in the configuration.

## Disclaimer

This automation script is an independent tool created to interface with Ellea1. It is not affiliated with or endorsed by Professor Eyal Levenberg or DTU.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

*Note: The Ellea1 Excel software is the intellectual property of Prof. Eyal Levenberg and is subject to its own license terms.*
