# AI-Powered Proposal Generation & Validation

## Overview

This project is part of the **AI-Powered Estimation Platform** and implements the **Proposal Generation Intelligence** capability.

The system automatically generates professional construction proposal documents from:
- Scope Intelligence output
- Cost Estimation output
  [output of Team 1 and 2] 

It also validates the generated proposal for completeness, consistency, contradictions, and readability.

## Problem Statement

Preparing construction proposals manually is time-consuming and prone to errors.

Estimators must:
- Review project scope
- Review cost estimates
- Write scope narratives
- Create exclusions and qualifications
- Prepare pricing summaries
- Verify proposal quality

This project automates proposal creation and validation, reducing manual effort and improving consistency.

## Folder Structure

```text
project/
│
├── main.py
│
├── modules/
│   ├── proposal_engine.py
│
├── config/
│   ├── settings.json
│   ├── template_schema.json
│   └── boilerplate.json
│
├── data/
│   ├── dummy_team1_output.json
│   └── dummy_team2_output.json
│
└── outputs/
    ├── generated_proposal.docx
    ├── completeness_report.txt
    ├── consistency_report.txt
    └── readability_report.txt
```

## Inputs
Team 1 Output contains Project information, Scope items, Scope gaps, Alternate options
Team 2 Output contains Cost estimate, Base bid total, Conflicts, Risk flags

---

## Outputs

### Proposal Document
```text
outputs/generated_proposal.docx
```

### Validation Reports
```text
outputs/completeness_report.txt
outputs/consistency_report.txt
outputs/readability_report.txt
```

---

## Requirements

- Python 3.10+
- Ollama
- Qwen 2.5 Models

Install dependencies:

```bash
pip install -r requirements.txt
```

Pull models:

```bash
ollama pull qwen2.5:7b
```

Start Ollama:

```bash
ollama serve
```

---

## Running the Project

Run with default sample inputs:

```bash
python main.py
```

Run with custom inputs:

```bash
python main.py --team1 data/team1.json --team2 data/team2.json
```

---

## Workflow

```text
Team 1 JSON
      +
Team 2 JSON
        ↓
Proposal Generation 
        ↓
Generated Proposal (.docx)
        ↓
Proposal Validation 
        ↓
Validation Reports
```
