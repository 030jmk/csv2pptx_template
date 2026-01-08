# csv2pptx_template
Generate PowerPoint presentations from Excel/CSV data. Each row becomes a slide with values mapped to named shapes in your template.

---

## Quick Start
```bash
# Setup (one time)
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Run
python slide_creator.py
```

---

## Setup

### 1. Create Virtual Environment
```bash
python3 -m venv venv
```

### 2. Activate Virtual Environment
```bash
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

---

## Usage

### Auto-detect files in current folder
```bash
python slide_creator.py
```

### Specify files manually
```bash
python gslide_creator.py -t template.pptx -d data.xlsx -o output.pptx
```

### Flags

| Flag | Description |
|------|-------------|
| `-t, --template` | Template .pptx file |
| `-d, --data` | Data file (.xlsx or .csv) |
| `-o, --output` | Output filename |

When auto-detecting, output filename includes timestamp: `output_20250108_143052.pptx`

---

## Template Setup

1. Create a PowerPoint file (**.pptx**, not .potx)
2. Design your slide with shapes (rectangles, text boxes, etc.)
3. Style each shape with desired font, size, color
4. Name shapes via: **Home → Arrange → Selection Pane**
5. Shape names must match Excel/CSV column headers exactly
6. Add placeholder text to each shape (don't leave empty)
7. Save as .pptx

> **Note:** SVG images won't copy to new slides. Use PNG or place SVGs in Slide Master.

---

## Data File Setup

- Excel (.xlsx) or CSV (.csv) supported
- First row = headers (must match shape names exactly)
- Each subsequent row = one slide
- CSV must be UTF-8 encoded

### Example

| DemoName | Date | Abstract |
|----------|------|----------|
| Demo One | 2025/Q1 | This is the abstract |
| Demo Two | 2025/Q2 | Another abstract |

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Text color wrong | Ensure shapes have placeholder text with correct styling |
| Only one slide generated | Check you're using .pptx not .potx |
| Shapes not populated | Verify shape names match headers exactly (Selection Pane) |
| "Multiple files found" error | Use `-t` and `-d` flags to specify files |

---

## Important Notes

- Shape names are **case-sensitive**
- Each shape needs placeholder text for formatting preservation
- Auto-detect requires exactly one .pptx and one .xlsx/.csv in folder
- Deactivate venv when done: `deactivate`
