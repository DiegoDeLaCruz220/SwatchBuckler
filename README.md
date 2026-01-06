# SwatchBuckler

A powerful color swatch extraction tool for designers and developers. Extract color swatches from product sheets, design documents, and images with ease.

## Features

- 🎨 **Automatic Color Detection** - Click on any color swatch to automatically detect its boundaries
- 📝 **OCR Text Recognition** - Automatically reads color names from labels using Tesseract OCR
- 🔍 **Zoom & Pan** - Smooth zooming and panning for precise selection
- 📁 **Custom Output Directory** - Choose where to save extracted swatches
- 🖱️ **Intuitive UI** - Easy-to-use interface with resizable panels
- ✅ **Selection Mode Toggle** - Switch between selection and navigation modes

## Installation

1. Clone this repository:
```bash
git clone https://github.com/DiegoDeLaCruz220/SwatchBuckler.git
cd SwatchBuckler
```

2. Install required dependencies:
```bash
pip install pillow pytesseract
```

3. Install Tesseract OCR:
   - **Windows**: Download and install from [UB-Mannheim Tesseract](https://github.com/UB-Mannheim/tesseract/wiki)
   - The script will automatically detect Tesseract in standard installation paths

## Usage

1. Run the application:
```bash
python extract_swatches_simple.py
```

2. Browse or paste the path to your image file

3. Click "Start" to open the image

4. **First time setup:**
   - Enable "Selection Mode"
   - Click on the first color swatch
   - Drag a box around the color's name/label
   - The app will learn the position pattern

5. **Extract remaining swatches:**
   - Click on each color swatch
   - The app will automatically detect the color name
   - Review/edit the detected name and press OK
   - All swatches are saved as PNG files

## Controls

- **Mouse Wheel** - Zoom in/out
- **Right-Click + Drag** - Pan around the image
- **F Key** - Fit image to window
- **Selection Toggle** - Enable/disable selection mode (prevents accidental selections while navigating)

## Output

All extracted swatches are saved as PNG files in the selected output directory with their color names as filenames (e.g., `dark_bronze.png`, `slate_blue.png`).

## Requirements

- Python 3.7+
- Pillow (PIL)
- pytesseract
- Tesseract OCR engine

## License

MIT License - feel free to use and modify as needed.

## Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.
