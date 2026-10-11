# Lab Report Generator

Single-file predecessor of LabScribe. Reads a Mindray BS240 analyser PDF and
produces a patient report in the Ngleshie Amanfrom Polyclinic layout.

    python lab_report.py --demo
    python lab_report.py --demo-hba1c
    python lab_report.py patient.pdf --path-no CHE901/26 --sex FEMALE --age 46

Needs Python 3.8+, `pdfplumber` and `reportlab`.

Kept for reference — LabScribe is the version in use. The demo patients are
invented; no real patient data belongs in this repository.
