# Auto-Generation of CANoe Configuration for CAN and LIN Networks

This repository provides a Python-based automation tool for generating Vector CANoe / CANalyzer configuration artifacts from CAN DBC and LIN LDF inputs, combined with structured Excel workbooks.

It represents the CAN and LIN network modules of a broader internship project focused on automating CANoe configuration generation for automotive system testing, conducted at Lear Corporation.

## Features

### CAN Network Support
- **DBC parsing**, Extracts CAN messages, signals, ECUs, and relevant metadata using a regex-based parser.

- **Excel integration**, Checksum initialization constants extracted from the `SIGNALS` sheet.

- **Artifact generation**
  - **Environment Variables (EV / ENV)**, Generation of environment variable declaration and value-description text files.
  - **CAPL scripts**
    - Standard CAPL generation
    - End-to-End (E2E) checksum variant
  - **CANoe panels**, XML-based `.xvp` panel definitions.

### LIN Network Support
- **LDF parsing**, Extracts LIN frames, signals, slave nodes, and relevant metadata using a regex-based parser.

- **Excel integration**, Checksum initialization constants extracted from slave-specific Excel sheets.

- **Artifact generation**
  - **Environment Variables (EV / ENV)**, Generation of environment variable declaration and value-description text files.
  - **CAPL scripts**, Standard CAPL generation.


## Project layout
- `src/canoe_can/can_config_generator.py`: CAN generator class.
- `src/canoe_lin/lin_config_generator.py`: LIN generator class.
- `scripts/demo_generate.py`: runnable example for CAN + LIN.
- `inputs/can/dbc/`: CAN DBC input file(s).
- `inputs/can/excel/`: CAN Excel input file(s).
- `inputs/can/img/`: logo images used in generated CAN panels.
- `inputs/lin/ldf/`: LIN LDF input file(s).
- `inputs/lin/excel/`: LIN Excel input file(s).
- `outputs/can/`: generated CAN artifacts (created at runtime).
- `outputs/lin/`: generated LIN artifacts (created at runtime).

## Usage
From the repo root:

Activate the virtual environment
```bash
source .venv/bin/activate
```

Run the demo generator
```bash
python3 -m scripts.demo_generate
```

The script:
- Locates the `.dbc` file under `inputs/can/dbc/` and the `*.xlsx` under `inputs/can/excel/`.
- Locates the `.ldf` file under `inputs/lin/ldf/` and the `*.xlsx` under `inputs/lin/excel/`.
- Generates output under `outputs/can/` and `outputs/lin/`.
