# Print Job Combo Optimizer

An automated production scheduling system that optimally groups print jobs for combined press runs using **Google OR-Tools CP-SAT**, reducing press setup time and material waste.

Built at **Bertelsmann SE** for book cover and jacket production scheduling across multiple US manufacturing locations.

---

## Impact

- Reduced manual planning effort by **50%**
- Runs automatically **twice daily** via Apache Airflow
- Fetches live data from **Snowflake**, exports color-coded Excel reports to **SharePoint**
- Handles **5,000+ print jobs** per run across 5 manufacturing locations

---

## Architecture

```
Snowflake (AI_COMBINED_RUN)
        ↓  fetch jobs (pandas + snowpark)
Data Preprocessing
        ↓  merge duplicate JOBs, parse ink/finish attributes
Group by (PRODUCTTYPE × PRESS_LOCATION × SEND_TO_LOCATION × PAPER)
        ↓
CP-SAT Optimizer (per group)
        ↓  bin-packing with hard + soft constraints
Color-coded Excel Report
        ↓
SharePoint Upload
        ↓
Apache Airflow (06:00 + 14:00 UTC)
```

---

## Optimization Rules

### Hard Constraints
| Rule | Detail |
|---|---|
| Capacity | Cover = max 4 jobs/combo, Jacket = max 2 jobs/combo |
| Sheet size | 28×20 paper → capacity reduced to 2 |
| Color compatibility | Special ink types must share identical signatures |
| Finishing conflicts | UV Overall ≠ Lamination, Gloss ≠ Matte lam, etc. |
| Quantity delta | Max ±3,500 units between jobs in same combo |
| 4/0 + specials | A 4/0 job limits specials in bin to max 1 |

### Soft Preferences (optimized)
- Minimize number of press runs (bins)
- Minimize quantity spread within combos
- Prefer 2× quantity ratios for covers (4-2-2 structure)
- Prefer pairing 4/4 jackets together

---

## Tech Stack

| Layer | Technology |
|---|---|
| Optimization | Google OR-Tools CP-SAT |
| Data warehouse | Snowflake (snowpark) |
| Orchestration | Apache Airflow |
| Data processing | pandas, numpy |
| Export | openpyxl (color-coded Excel) |
| SharePoint upload | office365-rest-python-client |
| Testing | pytest |

---

## Project Structure

```
print-job-optimizer/
├── src/
│   ├── models.py           # PrintJob dataclass
│   ├── finishing_rules.py  # Ink, UV, lam, foil, emboss compatibility rules
│   ├── optimizer.py        # CP-SAT solver
│   ├── data_loader.py      # Snowflake + CSV loader
│   └── export.py           # Color-coded Excel export
├── dags/
│   └── print_optimizer_dag.py  # Airflow DAG (runs twice daily)
├── data/
│   └── sample_jobs.csv     # Synthetic demo data (150 jobs)
├── tests/
│   └── test_optimizer.py   # Unit tests (pytest)
├── main.py                 # CLI entry point
├── .env.example
├── requirements.txt
└── README.md
```

---

## Setup & Run

### 1. Clone
```bash
git clone https://github.com/vaishnavi28-s/print-job-optimizer.git
cd print-job-optimizer
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run in demo mode (no credentials needed)
```bash
python main.py
```
Uses `data/sample_jobs.csv` with 150 synthetic print jobs.
Output: `combo_output.xlsx`

### 4. Run with Snowflake (production)
```bash
cp .env.example .env
# Fill in your Snowflake credentials
python main.py --snowflake
```

### 5. Run tests
```bash
pytest tests/ -v
```

---

## Output

The optimizer produces a color-coded Excel workbook with one tab per manufacturing location:

| Tab | Content |
|---|---|
| `<location>` | Combined jobs (2+ per combo), color-coded by ComboID |
| `<location>_SINGLES` | Jobs that could not be combined |

Each color block = one press run. Jobs in the same color share compatible finishing, ink type, paper, and quantity range.

---

## Sample Data

`data/sample_jobs.csv` contains 150 synthetic print jobs across 5 locations with realistic:
- Product types (Cover, Jacket)
- Paper specifications (28×20, 25×38, etc.)
- Ink configurations (4/0, 4/1, 5/0, 4/4, special inks)
- Finishing operations (lamination types, UV, foil, emboss)
- Quantity ranges (1,000 – 20,000 units)

> Production version connects to Snowflake and processes 5,000+ live jobs per run.

---

## Note on Data

This repository uses **synthetic demo data only**. No proprietary Bertelsmann data is included. Production credentials and real job data are managed via environment variables and are never committed to version control.

---

## Connect
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue)](https://www.linkedin.com/in/vaishnavi-sreekumar-48199a197/)
[![GitHub](https://img.shields.io/badge/GitHub-Follow-black)](https://github.com/vaishnavi28-s)
