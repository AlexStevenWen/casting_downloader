# access_downloader

**access_downloader** is an automated database and CAD file acquisition tool designed to build a synchronized bridge between structured attribute data and unstructured 3D geometric files. It serves as a core data-preprocessing component for the master's thesis research on **Classification and Process Prediction of Gating Systems for Investment Casting of A356 Aluminum Alloy**. By reading JSON configuration files, the tool automatically exports Microsoft Access tables and precisely locates and copies associated CAD (STEP) files.

## Project Background and Academic Context

This tool is part of the experimental methodology of the following master's thesis:

- **Thesis Title**: A Study on the Classification and Process Prediction of Gating Systems for Investment Casting of A356 Aluminum Alloy
- **Author**: HSU, WEN-HO
- **Advisor**: CHEN, TZUNG-MING
- **Institution**: National Changhua University of Education
- **Department**: Department of Electrical and Mechanical Technology
- **Degree**: Master's Thesis
- **Oral Defense Date**: 2026-07-10
- **Permanent URL**: [https://hdl.handle.net/11296/32a644](https://hdl.handle.net/11296/32a644)

**Keywords**: A356 aluminum alloy, investment casting, automated machine learning (AutoML), 3D point cloud semantic segmentation.

The research aims to establish a gating system classification model and process prediction system for A356 aluminum alloy investment casting. To achieve this, it is necessary to integrate process parameters (structured data) and 3D geometric models (unstructured data) from a large number of castings, supporting subsequent automated machine learning (AutoML) modeling and 3D point cloud semantic segmentation analysis. `access_downloader` was developed as a key tool for this data preparation stage, responsible for automatically extracting, pairing, and synchronizing the required training datasets from the original database and file system.

## Key Features

*   **Dual-Mode Operation Strategy**: Supports a "Full Backup Mode" (unconditionally parsing all data and globally scanning/copying CAD files) and a "Target Feature Extraction Mode" (precisely retrieving associated data and files based on a provided list of specific product versions).
*   **Data and Geometric Topology Synchronization**: When retrieving process parameters, relation-aware synchronization logic automatically locates the corresponding "blank" and "assembly" CAD files for the specified version and bundles them together, fundamentally eliminating the issue of missing geometric references.
*   **Automated Environment Detection**: Built-in scripts automatically scan system drives (e.g., `C:\`, `D:\Program Files`) to locate the FreeCAD execution environment and specific project directories (e.g., `ChynWangProject`, `ChynWangData`).
*   **Support for AutoML and Point Cloud Segmentation Data Flow**: The synchronized datasets output by the tool can be directly used as input for subsequent automated machine learning modeling and 3D point cloud semantic segmentation.

## Installation

This project uses Anaconda for environment management. Please ensure that Anaconda or Miniconda is installed on your system.

1. Create the virtual environment and install dependencies:
```bash
conda env create -f environment.yml
Activate the environment:

bash
conda activate access_downloder
Note: The environment name follows the original project configuration. If you rename it in environment.yml, use the corresponding name.

Usage
The program is executed via the command-line interface (CLI) and requires a JSON configuration file. This design deeply decouples the core computational logic from the environment configuration, ensuring that no underlying architecture changes are needed when switching environments.

Execution Command

bash
python main.py --config config.json
Configuration Example (config.json)

json
{
    "db_path": "C:\\path\\to\\your\\database.accdb",
    "table_export_path": "C:\\path\\to\\export\\tables",
    "cad_source_path": "C:\\path\\to\\source\\CAD\\files",
    "cad_target_path": "C:\\path\\to\\target\\CAD\\files",
    "product_versions": ["Version_A", "Version_B"]
}
Parameter Notes:

db_path, table_export_path, cad_source_path, and cad_target_path are required.

product_versions is optional. If this list is not provided, the system enters "Full Backup Mode." If a list of specific versions is provided, the system enters "Target Feature Extraction Mode."

Core Dependencies
This environment is based on python=3.11.13. Key packages include:

Data Processing: pandas, numpy, openpyxl

Database Connectivity: pyodbc, sqlite

System & Utilities: tqdm, colorama

Exit Codes & Output
Upon completion, the program returns a structured JSON result indicating the status and terminates with the corresponding exit code:

Exit Code 0 (status: success): All database exports and CAD file copying tasks completed successfully.

Exit Code 1 (status: error): Failed to read the configuration file, missing required parameters, or an error occurred during task execution.

Citation
If this tool is helpful to your research, please cite the following thesis:

HSU, W.-H. (2026). A Study on the Classification and Process Prediction of Gating Systems for Investment Casting of A356 Aluminum Alloy (Master's thesis). National Changhua University of Education, Department of Electrical and Mechanical Technology, Changhua City. Retrieved from https://hdl.handle.net/11296/32a644
