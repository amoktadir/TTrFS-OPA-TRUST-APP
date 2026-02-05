# Integrated TTrFS-OPA-TRUST MCDM Models (Streamlit App)

A Streamlit web application that implements **two Multi-Criteria Decision-Making (MCDM) models**:

1. **Trigonometric Trapezoidal Fuzzy OPA (TTrF-OPA)** — multi-expert criteria weighting using linguistic assessments aggregated with a trigonometric trapezoidal fuzzy weighted geometric operator, then solved via a trapezoidal linear programming model.
2. **TTrFS-TRUST Method** — a **multi-normalization, multi-distance** assessment framework that supports **Soft (linguistic)** and **Hard (crisp)** criteria, expert aggregation, constraint-based normalization, and ranking of alternatives.

The app also supports **exporting results to Word (.docx)** for both models.

---

## ✨ Features

### ✅ Trigonometric Trapezoidal Fuzzy OPA (TTrF-OPA)
- Define **criteria** and **number of experts**
- Assign **expert weights** (validated to sum to 1)
- Experts provide **linguistic ratings** (ELI … EHI)
- Aggregates fuzzy importance via **TTrFWG** (trigonometric trapezoidal fuzzy weighted geometric)
- Builds **OPA coefficient set (θ)** and solves **fuzzy LP** using **PuLP**
- Displays:
  - aggregated fuzzy importance (θ)
  - coefficients
  - fuzzy weights (l, m, u, w)
  - ranked criteria
  - Ψ (psi) and defuzzified ψ
- **Export**: Word report of all key tables

### ✅ TTrFS-TRUST Method
- Step-by-step workflow with progress navigation:
  1. Problem Setup (alternatives, criteria, experts, α, β)
  2. Criteria Setup (Soft/Hard)
  3. Expert Weights
  4. Data Collection (soft linguistic + hard crisp)
  5. Decision Matrix (aggregation + defuzzification)
  6. Criteria Info (Benefit/Cost + weights)
  7. Constraint Intervals (ρᴸ, ρᵁ)
  8. Results (multi-normalization + multi-distance ranking)
- Supports **four normalization schemes**:
  - Linear ratio-based (r)
  - Linear sum-based (s)
  - Max–Min (m)
  - Logarithmic (l)
- Computes distance measures:
  - Euclidean (ε)
  - Manhattan (π)
  - Lorentzian (ℓ)
  - Pearson (ρ)
- Produces final score **ℒ** and ranking
- **Export**: Word report including decision matrix, normalization matrices, and final ranking

---

## 🧰 Tech Stack

- **Python**
- **Streamlit** (UI)
- **NumPy / Pandas** (data & computation)
- **PuLP** (linear programming)
- **python-docx** (Word report generation)

---

## 📦 Project Structure (Recommended)

├─ TTrFS-OPA-TRUST.py                 # main Streamlit app (code)
├─ requirements.txt
├─ README.md
└─ assets/                # optional screenshots / demo gifs


## 🧪 Usage Guide

### A) TTrF-OPA Model (Criteria Weighting)
1. Select **“Trigonometric Trapezoidal Fuzzy OPA”** from the sidebar.
2. Enter:
   - Number of experts
   - Number of criteria
   - Criterion names
3. Set **expert weights** (must sum to **1.00**).
4. For each expert, choose a linguistic rating for each criterion:
   - `ELI, VLI, LI, MI, HI, VHI, EHI`
5. Click **Calculate Weights**.
6. Review:
   - aggregated θ
   - coefficients
   - fuzzy weights and rank
   - ψ values
7. Click **Export Results to Word** to download a report.

### B) TTrFS-TRUST Method (Ranking Alternatives)
1. Select **“TTrFS-TRUST Method”** from the sidebar.
2. Follow the steps in order:
   - **Problem Setup:** set α (must sum to 1) and β
   - **Criteria Setup:** mark each criterion as Soft/Hard
   - **Expert Weights:** must sum to 1
   - **Data Collection:**
     - Soft criteria → linguistic ratings by each expert
     - Hard criteria → crisp numeric values
   - **Decision Matrix:** aggregated and defuzzified matrix
   - **Criteria Information:** Benefit/Cost + criterion weights (sum to 1)
   - **Constraints:** enter ρᴸ and ρᵁ (bounds)
   - **Results:** view normalization matrices, distances, and final ℒ ranking
3. Export the report using **Export TRUST Results to Word**.

---

## 🧠 Method Notes (Quick)

### Linguistic Terms
- OPA model uses:
  - `ELI, VLI, LI, MI, HI, VHI, EHI`
- TRUST model uses:
  - `ELI, VLI, LI, MLI, MI, MHI, HI, VHI, EHI`

### Key Parameters
- **α = (∂₁, ∂₂, ∂₃, ∂₄)**: weights for the four normalization methods (should sum to 1)
- **β**: distance aggregation parameter (0 to 1)

---

## ☁️ Deployment

### Streamlit Community Cloud
1. Push this repo to GitHub.
2. Go to Streamlit Community Cloud and select the repository.
3. Set:
   - **Main file path:** `TTrFS-OPA-TRUST.py`
4. Ensure `requirements.txt` is present.

---
## 📄 Citation

If you use this app in academic work, please cite:

**Moktadir, M. A., Lu, S., & Ren, J. (2026). A Novel Environmental, Social, and Governance Performance Assessment Model for the Global Oil and Gas Sector. Business Strategy and the Environment, doi: https://doi.org/10.1002/bse.70614.**


---

## 👤 Author

Developed by **Md Abdul Moktadir**

---
