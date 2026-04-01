# 📈 Excel VBA & Sheet Extractor

A lightweight, containerized web utility built on **Alpine Linux** to quickly extract VBA macro code and convert Excel worksheets into CSV format. 

Everything is bundled into a single ZIP for easy local download. No Excel installation required.

## ✨ Features
* **VBA Extraction:** Pulls all `.vb` macro code from `.xls`, `.xlsm`, and `.xlsb` files.
* **Sheet Conversion:** Automatically exports every worksheet into an individual `.csv` file.
* **Privacy-First:** Processes everything locally within your container.
* **Ultra-Lightweight:** Uses an Alpine-based Python image to minimize the footprint.
* **Drag & Drop:** Simple, modern web interface.

---

## 🚀 Getting Started

### Prerequisites
* [Podman](https://podman.io/) (or Docker)

### Installation & Run

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/excel-extractor.git
   cd excel-extractor
   ```

2. **Build the image:**
   ```bash
   podman build -t excel-extractor .
   ```

3. **Launch the container:**
   ```bash
   podman run -d \
     --name excel-tool \
     -p 5000:5000 \
     excel-extractor
   ```

4. **Access the app:**
   Open your browser and navigate to `http://localhost:5000`.

---

## 🛠️ Tech Stack
* **Base OS:** Alpine Linux
* **Backend:** Python 3.11 / Flask
* **Parsing:** * `oletools`: For robust VBA stream extraction.
    * `pandas` & `openpyxl`: For high-speed worksheet processing.
* **Frontend:** Vanilla JS (Drag & Drop API)

---

## 📝 Usage Note
The extractor handles the two most common Excel formats:
1.  **.xlsx / .xlsm**: Modern XML-based formats.
2.  **.xls**: Legacy OLE formats.

When you upload a file, the tool generates a ZIP containing:
* `/vba/`: All detected macros as `.vb` files.
* `/sheets/`: All worksheets as `.csv` files.

---

## ⚖️ License
Distributed under the MIT License. See `LICENSE` for more information.

---

**Tip:** Since this runs on Alpine, the image size is kept small. If you're running this in a production-like environment, you may want to swap the Flask dev server for `gunicorn` in the `Dockerfile`.