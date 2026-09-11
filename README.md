README.md
Markdown
# access_downloader

**access_downloader** is an automated database and CAD file acquisition tool designed to build a synchronized bridge between structured attribute data and unstructured 3D geometric files. Tailored for the A356 aluminum alloy investment casting prediction system, this tool reads JSON configuration files to automatically export Microsoft Access tables and precisely locate and copy associated CAD (STEP) files.

## Key Features

*   **Dual-Mode Operation Strategy**: Supports a "Full Backup Mode" (unconditionally parsing all data and globally scanning/copying CAD files) and a "Target Feature Extraction Mode" (precisely retrieving associated data and files based on a provided list of specific product versions).
*   **Data and Geometric Topology Synchronization**: When retrieving process parameters, a relation-aware synchronization logic automatically locates the corresponding "blank" (毛胚) and "assembly" (組樹) CAD files for the specified version and bundles them together, fundamentally eliminating the issue of missing geometric references[cite: 1].
*   **Automated Environment Detection**: Built-in scripts automatically scan system drives (e.g., `C:\`, `D:\Program Files`) to locate the FreeCAD execution environment and specific project directories (e.g., `ChynWangProject`, `ChynWangData`).

## Installation

This project uses Anaconda for environment management. Please ensure that Anaconda or Miniconda is installed on your system.

1. Create the virtual environment and install dependencies:
```bash
conda env create -f environment.yml
Activate the environment:

Bash
conda activate access_downloder
Usage
The program is executed via the command line interface (CLI) and requires a JSON configuration file. This design deeply decouples the core computational logic from the environment configuration, ensuring that no underlying architecture changes are needed when switching environments[cite: 1].

Execution Command
Bash
python main.py --config config.json
Configuration Example (config.json)
JSON
{
    "db_path": "C:\\path\\to\\your\\database.accdb",
    "table_export_path": "C:\\path\\to\\export\\tables",
    "cad_source_path": "C:\\path\\to\\source\\CAD\\files",
    "cad_target_path": "C:\\path\\to\\target\\CAD\\files",
    "product_versions": ["Version_A", "Version_B"] 
}
Parameter Notes:

db_path, table_export_path, cad_source_path, and cad_target_path are required.

product_versions is optional. If this list is not provided, the system will enter "Full Backup Mode"[cite: 1]. If a list of specific versions is provided, the system will enter "Target Feature Extraction Mode"[cite: 1].

Core Dependencies
This environment is based on python=3.11.13. Key packages include:

Data Processing: pandas, numpy, openpyxl

Database Connectivity: pyodbc, sqlite

System & Utilities: tqdm, colorama

Exit Codes & Output
Upon completion, the program returns a structured JSON result indicating the status and terminates with the corresponding exit code:

Exit Code 0 (status: success): All database exports and CAD file copying tasks completed successfully.

Exit Code 1 (status: error): Failed to read the configuration file, missing required parameters, or an error occurred during task execution.
