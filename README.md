# CST to CSV Export Tool

A Python-based GUI application for exporting CST Studio Suite S-parameter simulation results to CSV and Excel formats.

## Features

- Export S11 parameters from CST files to CSV or Excel format
- Multiple display modes:
  - Complex (Real + Imaginary)
  - Magnitude
  - Magnitude (dB)
  - Magnitude and Phase
- Frequency range selection
- Parameter sweep support
- Summary view showing:
  - Frequency range
  - Number of frequency points
  - Parameter combinations
  - Minimum S11 value with corresponding frequency
- Column selection for export
- Modern, user-friendly interface

## Installation

### Option 1: Standalone Executable (Windows)

1. Download the latest release from the [Releases](../../releases) page
2. Extract the ZIP file
3. Run `cst2csv.exe`
4. On first run, select your CST Studio Suite Python libraries path (typically `C:\Program Files (x86)\CST Studio Suite 20XX\AMD64\python_cst_libraries`)

### Option 2: From Source

1. Clone the repository:
```bash
git clone https://github.com/ismailakdag/cst2csvgui2024.git
cd cst2csvgui2024
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the application:
```bash
python cst2csv.py
```

## Requirements

- Python 3.11 or higher
- CST Studio Suite 2024 (or compatible version)
- Required Python packages (automatically installed):
  - PyQt5
  - pandas
  - numpy
  - openpyxl

## Usage

1. Launch the application
2. If first time, select your CST Studio Suite Python libraries path
3. Click "Browse" to select your .cst file
4. Choose display mode (Complex, Magnitude, dB, or Phase)
5. Optionally adjust frequency range
6. Select columns to export
7. Click "Export..." and choose save location

## Configuration

- CST library path can be changed via Settings → Change CST Library Path
- Configuration is stored in `config.json` in the application directory

## Known Limitations

- Currently supports S11 parameters only
- Requires CST Studio Suite Python libraries
- Windows OS support only

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Acknowledgments

- CST Studio Suite for providing the Python libraries
- PyQt5 for the GUI framework
