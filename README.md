```markdown
📂 PDF2PACS  

**PDF2PACS** is a lightweight desktop application that allows healthcare professionals to easily upload PDF documents into a PACS archive.  
It encapsulates PDFs as DICOM objects and sends them via C-STORE to the configured PACS.  

---

📖 Extended Description  

🩺 For End Users (Radiologists, Technicians, Clinicians)  
**PDF2PACS** is a user-friendly desktop application designed to simplify the integration of non-DICOM documents into radiology workflows.  
With just a few clicks, healthcare professionals can upload PDF reports, scanned papers, or administrative documents directly into the PACS archive.  

The application automatically encapsulates PDF files into the DICOM format, so they can be stored and viewed in the same system as imaging studies. Supporting both single and multiple file uploads, **PDF2PACS** ensures that external reports are seamlessly attached to the patient’s imaging record, keeping all relevant information centralized and accessible within the PACS.  

  Why it’s useful:**  
- Centralizes all patient-related data (images + documents)  
- Simplifies workflows for radiology staff  
- Reduces reliance on paper archives and external storage  
- Ensures consistency and accessibility of patient history  

---

 💻 For Developers & Integrators  
**PDF2PACS** is implemented in Python and simulates a DICOM modality to transfer PDF files as *Encapsulated PDF* objects using the DICOM C-STORE protocol.  

**Technical highlights:**  
- **Language:** Python 3  
- **Dependencies:** `pydicom`, `requests`, `tkinter` (UI), optional packaging with `pyinstaller`  
- **Workflow:**  
  1. Select one or more PDF files through the GUI  
  2. Each PDF is wrapped into a valid DICOM object with configurable metadata (Patient Name, ID, Study Description, Series Description)  
  3. Objects are sent to the target PACS via C-STORE SCU  
- **Configurable settings:**  
  - PACS IP, Port, and AE Title  
  - Default DICOM metadata values  
- **Extensible:** Developers can easily adapt the code to handle other document types (e.g., JPEG, DOCX → encapsulated as DICOM) or integrate with custom hospital systems  

This makes the project suitable both as a production tool in small/medium radiology departments and as a learning resource for developers exploring DICOM communication and PACS integration.  

---

## 🖼️ Screenshot  
![PDF2PACS](https://github.com/Algernon72/PDF2PACS/blob/main/PDF2PACS.png?raw=true)



## ⚙️ Requirements  

- Python **3.9+**  
- PACS or Orthanc server reachable via DICOM (C-STORE enabled)  
- Windows, macOS, or Linux  

**Python dependencies:**  
```

pydicom
requests
tkinter

````

---

🚀 Installation  

Clone the repository and install dependencies:  

```bash
git clone https://github.com/your-username/pdf2pacs.git
cd pdf2pacs
pip install -r requirements.txt
````

*(If tkinter is not included in your Python distribution, install it separately according to your OS.)*

---

▶️ Usage

Run the application with:

```bash
python Modality_pdf_uploader.py
```

Steps:

1. Launch the application
2. Configure PACS settings (IP, Port, AE Title) in the settings menu
3. Select one or more PDF files
4. Click **Invia a Orthanc** to encapsulate and send to PACS
---

````markdown
## 🛠️ Build Instructions  

The project includes a ready-to-use **`.spec`** file and a **`.bat`** script to simplify the build process with PyInstaller.  
These allow you to generate a smaller executable that depends on external files, instead of a large one-file binary.  

### 🔹 Build with the provided `.bat` script  
Simply run:  

```bash
build_small_exe.bat
````

This will call PyInstaller using the two provided `.spec` files and create the executables inside the `dist/` folder.

### 📂 Output

* The compiled **exe** will be created in the `dist/` directory
* The required resource files will be placed in the same folder structure defined in the `.spec` file
* Keep the `.exe` and the resource files together in the same directory for the application to run correctly

---

⚠️ **Note:**

* The `.bat` script assumes that Python and PyInstaller are installed and available in your PATH
* If you prefer a fully self-contained executable (larger size), you can still use:

```bash
pyinstaller --onefile --noconsole pdf2pacs.py
```

```

---

Vuoi che ti prepari anche un esempio di **build.bat** e di **pdf2pacs.spec** minimale da mettere nel repo, così il README è subito coerente con i file?
```

---

## 📜 License

This project is released under the **MIT License**. You are free to use, modify, and distribute it with attribution.

---

## 🤝 Contributing

Contributions are welcome!

* Fork the repository
* Create a feature branch (`git checkout -b feature-name`)
* Commit your changes and push
* Open a Pull Request

---

## 📧 Contact

For questions, suggestions, or bug reports, feel free to open an issue on GitHub or contact the maintainer directly.
