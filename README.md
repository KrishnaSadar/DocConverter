# DOCX Question Bank Converter (GitHub Pages Ready)

A fast, client-side static web application designed to convert raw question papers (`input_format.docx`) into the strict 10-row system-ready table layout (`output_format.docx`).

## Features

- **100% Client-Side Conversion**: Runs entirely in the user's browser using JavaScript and JSZip. No data or documents are ever uploaded to an external server.
- **Embedded Media & Equation Preservation**: Transfers preview images and ChemDraw / MS Equation OLE `.bin` objects and remaps all relationship IDs seamlessly.
- **Answer Key & Solution Mapping**: Automatically extracts answers and hints/solutions from the document tables and links them to the corresponding questions.
- **Built-in Output Template**: Comes pre-packaged with the official `output_format.docx` template base, eliminating the need to upload a template file each time.
- **Verification Dashboard**: Live, searchable in-browser preview of all converted questions, options, answers, and solutions before downloading.
- **GitHub Pages Ready**: Zero build steps or backend servers needed. Simply push to GitHub and enable Pages in repository settings.

## File Structure

```text
├── index.html          # Main application page
├── style.css           # Modern dark-mode styling with glassmorphism
├── converter.js        # Core DOCX parsing, relationship remapping, and generation engine
├── app.js              # UI controller, drag & drop, preview, and download handler
├── template-data.js    # Pre-packaged base64 template from output_format.docx
├── lib/
│   └── jszip.min.js    # Bundled JSZip library for 100% offline support
├── input_format.docx   # Reference input document (used for the "Try Sample" demo)
└── output_format.docx  # Reference output document
```

## How to Deploy to GitHub Pages

1. Push these files to your GitHub repository.
2. Go to **Settings** > **Pages** in your repository.
3. Under **Branch**, select `main` (or `master`) and folder `/ (root)`.
4. Click **Save**. Your converter website will be live at `https://<username>.github.io/<repo>/`!

## Local Usage

Open `index.html` directly in your browser, or start any local server:
```bash
python -m http.server 8080
```
Then visit `http://localhost:8080` in your web browser.
