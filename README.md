
# CapSlide

**CapSlide** is a streamlined Command Line Interface (CLI) tool designed to automate the conversion of subtitle files (`.json` or `.txt`) into professional PowerPoint (PPTX) presentations. By utilizing customizable templates and smart placeholders, it eliminates the manual effort of copy-pasting content for lectures, presentations, or video summaries.

---

## ✨ Key Features

* **Multi-format Support**: Seamlessly parse both `.json` and `.txt` subtitle formats.
* **Template-Driven**: Use your own `.pptx` files as templates to maintain consistent branding and design.
* **Precision Positioning**: Specify the exact slide index within your template for content injection.
* **Text Sanitization**: Optional punctuation filtering (`--ignore_marks`) for cleaner, more professional slides.
* **Developer Friendly**: Built-in verbose logging for easy debugging and process tracking.

---

## 🚀 Quick Start

### 1. Prerequisites
* Python 3.7 or higher
* `python-pptx` library

### 2. Installation
Clone the repository and install the required dependencies:

```bash
git clone [https://github.com/Annabelle1899/CapSlide.git](https://github.com/Annabelle1899/CapSlide.git)
cd CapSlide
pip install -r requirements.txt

```

### 3. Usage

Run the tool via the command line as follows:

```bash
python -m capslide.main [INPUT_FILE] --template [TEMPLATE_PATH] [OPTIONS]

```

#### Arguments & Options:

| Argument | Shorthand | Description | Required | Default |
| --- | --- | --- | --- | --- |
| `input` | N/A | Path to the source subtitle file (.json/.txt). | **Yes** | - |
| `--template` | `-t` | Path to the template PPTX file. | **Yes** | - |
| `--output` | `-o` | Path for the generated PPTX file. | No | `output.pptx` |
| `--placeholder` | `-p` | The specific text tag in the template to be replaced. | No | `subtitle` |
| `--template_slide_page_number` | `-n` | Slide index (0-based). Use `-1` for the last slide. | No | `-1` |
| `--ignore_marks` | `-i` | Exclude punctuation marks from the slides. | No | False |
| `--verbose` | `-v` | Display detailed processing logs in the console. | No | False |

### 4. Example

Generate a presentation using a specific slide from a template and ignoring punctuation:

```bash
python -m capslide.main lecture.txt -t assets/modern_style.pptx -n 2 -i -o final_presentation.pptx

```

---

## 📁 Project Structure

```text
CapSlide/
├── capslide/
│   ├── __init__.py
│   ├── main.py          # CLI Entry point & Argument Parsing
│   └── core.py          # Core processing logic & SubtitlesProcessor
├── requirements.txt     # Project dependencies (e.g., python-pptx)
└── README.md            # Project documentation

```

---

## 🛠 Contributing

Contributions make the open-source community an amazing place to learn and create.

1. **Fork** the Project.
2. **Create** your Feature Branch (`git checkout -b feature/AmazingFeature`).
3. **Commit** your Changes (`git commit -m 'Add some AmazingFeature'`).
4. **Push** to the Branch (`git push origin feature/AmazingFeature`).
5. **Open** a Pull Request.

---

## 📄 License

Distributed under the MIT License. See `LICENSE` for more information.

---

**CapSlide** — Transforming subtitles into presentations with ease.

