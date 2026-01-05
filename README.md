# Auto-Generation of CANoe Configuration for CAN Networks

This repository presents a Python-based automation tool that generates Vector CANoe / CANalyzer configuration artifacts from a CAN DBC file and a structured Excel workbook.

This repository represents the **CAN network module** of a broader internship project focused on the automated generation of CANoe configurations for automotive system testing across **CAN and LIN networks** at Lear Corporation.

## Features

- **DBC parsing**: Extracts CAN messages, signals, nodes, and metadata using a regex-based parser.

- **Excel integration**: Reads checksum initialization constants from an Excel workbook (SIGNALS sheet).

- **Artifact generation**:

  - **Environment Variables**: EV/ENV declaration and value description text files.

  - **CAPL scripts**: 
    - Standard CAPL generation
    - End-to-End (E2E) checksum variant
 
  - **CANoe panels**: XML-based `.xvp` panel definitions

## Project layout
- `src/canoe_autogen_can/can_canoe_config.py`: monolithic generator class.
- `scripts/demo_generate.py`: runnable example.
- `inputs/dbc/`: DBC input file(s).
- `inputs/excel/`: Excel input file(s).
- `inputs/img/`: logo images used in generated panels.
- `outputs/`: generated artifacts (created at runtime).

## Usage
From the repo root:

```bash
python scripts/demo_generate.py
```

The script:
- Locates the `.dbc` file under `inputs/dbc/` and the `*.xlsx` under `inputs/excel/`.
- Generates output under `outputs/`.


