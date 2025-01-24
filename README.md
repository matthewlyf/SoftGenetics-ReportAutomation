# Automated Folder Monitoring and Patient Report Generator

This Python script was developed for **Hamilton Health Sciences (HHSC)** to streamline the process of generating patient-specific genomic reports from data processed by SoftGenetics Geneticist Assistant. It is highly customized for the workflows, file structures, and reporting requirements of HHSC and may have limited applicability outside this context.  It scans a queue folder for a text file containing the directories for which a report needs to be generated and triggers the reporting script.

---

## Features

- **Custom Workflow Integration**:
  - Monitors a designated folder for `.txt` files that trigger report generation.
  - Reads SoftGenetics Geneticist Assistant data to generate detailed patient reports.
- **Automated Reporting**:
  - Processes genomic data, including variant coverage and classification.
  - Generates patient-specific Word reports using predefined templates.
- **Error Handling**:
  - Validates input file paths and ensures invalid or incomplete data is flagged.
- **Comprehensive Data Analysis**:
  - Filters and classifies variants (e.g., Tier I and Tier II).
  - Incorporates functional evidence and COSMIC database annotations.

---

## Why It’s Specialized for HHSC

This script is designed to fit seamlessly into the workflow of HHSC, addressing their specific needs:
- **Data Sources**: Tailored to parse file structures and outputs from SoftGenetics Geneticist Assistant.
- **Report Templates**: Uses pre-configured Word document templates that match HHSC's reporting standards.
- **Highly Customized Logic**: Includes institution-specific classification criteria and functional evidence references.

Due to these customizations, the script is unlikely to be directly applicable to other institutions without significant modifications.

---

## Requirements

- Python 3.8+
- Libraries:
  - `os`, `time`, `shutil`, `pandas`, `docx`, `requests`, `re`, `decimal`

Install dependencies using:

```bash
pip install pandas python-docx requests
```

---

## Usage

1. **Set Up Directories**:
   - Define the monitored folder (`folder_directory`) and the subfolder for completed files (`completed_directory`).

2. **Configure the Script**:
   - Replace `[PATH_TO_WATCHED_FOLDER]`, `[PATH_TO_REPORT_TEMPLATE]`, and `[PATH_TO_GENE_STATEMENTS]` with the appropriate HHSC paths.

3. **Run the Script**:
   ```bash
   python auto_reporter.py
   ```

4. **Place `.txt` Files**:
   - Add `.txt` files to the monitored folder, each containing a path to a directory with the necessary genomic data.

5. **View Reports**:
   - Reports are generated and saved in the corresponding patient data directory.

---

## Example Workflow

1. A `.txt` file named `patient_001.txt` is placed in the monitored folder, containing:
   ```
   /path/to/patient_001_data
   ```

2. The script:
   - Reads data from `/path/to/patient_001_data`.
   - Processes the coverage and variant information.
   - Generates a Word report tailored to the patient, saved in `/path/to/patient_001_data`.

3. The `.txt` file is moved to the `completed` folder for tracking.

---

## Limitations

- **Institution-Specific**: Designed exclusively for HHSC’s genomic analysis workflow.
- **Rigid File Naming**: Requires adherence to specific file and folder naming conventions.
- **Custom Templates**: Report templates are unique to HHSC and need significant reconfiguration for other use cases.

---

## Future Enhancements

- Integration with additional genomic analysis tools.
- Improved error handling for unsupported file formats.
- Options for real-time notifications upon report completion.

---


