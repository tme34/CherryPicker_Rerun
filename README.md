# Biomek Genotyping Rerun & Verification Suite

<p align="center">
  <img src="https://img.shields.io/badge/Liquid%20Handling-Beckman%20Coulter-blue" alt="Biomek">
  <img src="https://img.shields.io/badge/Workflow-Automation-green" alt="Automation">
  <img src="https://img.shields.io/badge/R-Data%20Processing-yellow" alt="R">
</p>

## 📋 Overview
This repository contains a specialized automation pipeline for managing **genotyping reruns** on Biomek liquid handling platforms (Span-8). 

When samples return "Undetermined" results, they must be "cherry-picked" from their original plates and consolidated into new plates for re-analysis. This suite automates the data-heavy task of generating instruction files for the robot and performing post-run verification to ensure sample integrity.

---

## 🛠 Script Details

### 1. Rerun_script_V2
The primary data processor. It transforms manual failure logs into a format the Biomek software can interpret.

* **Logic:** It reads a list of failed samples and maps them to their physical locations.
* **Worklists:** Instructions for the Biomek to pick specific samples.
* **Reverse Worklists:** Essential for data traceability; these files map the **Destination** wells back to the **Source** wells so that once the rerun is finished, the results can be correctly re-attached to the original Sample IDs.
* **Gathered Worklist:** A master Excel summary for lab documentation and high-level review.

### 2. Destination_Scan_Tjek_script
The secondary verification script (Quality Control).

* **Logic:** This script acts as a "digital handshake." It compares the **Worklist** (the digital plan) against a **Scan File** (the physical reality).
* **Validation:** It verifies that the `WellBarcode` and `ID` scanned from the final 96-well destination plate match the intended output. This prevents errors caused by accidental plate rotation or incorrect source plate loading.

---

## 🧬 Laboratory Workflow

1.  **Identify Failures:** Generate an Excel file of "Undetermined" samples.
2.  **Generate Files:** Run `Rerun_script_V2` to create the `/Worklists/` and `/Reverse_Worklists/` folders.
3.  **Robot Execution:** Load the worklist into the Biomek Method. The Span-8 pod picks samples from the **Master Plates** into the **96-well Rerun Plate**.
4.  **Verification Scan:** Scan the final 96-well plate.
5.  **Run Tjek:** Use `Destination_Scan_Tjek_script` to compare the scan against the worklist. 



---

## 📊 Data Input Format
The scripts expect an Excel file with the following columns, but the headers should NOT be present:

| Header | Description |
| :--- | :--- |
| **ID** | Unique identifier for the sample |
| **MasterPlate** | Name/ID of the source plate |
| **Position** | Well coordinate (e.g., A1, C10) |
| **PlateBarcode** | Physical barcode of the source plate |
| **WellBarcode** | Unique barcode for the specific tube/well |

---

## 📁 File Structure
```text
├── Rerun_script_V2.py              # Main worklist generator
├── Destination_Scan_Tjek_script.py # Verification script
├── rerunbatchV2.bat                # Windows-batchfile
├── Dest_scan_batch.bat             # Windows-batchfile
├── /Worklists/                     # Output for Biomek instructions
└── /Reverse_Worklists/             # Output for traceability mapping




