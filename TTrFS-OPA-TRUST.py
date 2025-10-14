import streamlit as st
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import base64
from pulp import LpMaximize, LpProblem, LpVariable, lpSum, value
import math

# Set page configuration
st.set_page_config(
    page_title="Integrated MCDM Models",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
st.markdown("""
<style>
:root {
    --primary: #1f77b4;
    --secondary: #2ca02c;
    --accent: #ff6b6b;
    --background: #f8f9fa;
    --card-bg: #ffffff;
    --text: #262730;
    --border-radius: 12px;
    --shadow: 0 6px 16px rgba(0, 0, 0, 0.08);
}

.main-header {
    font-size: 2.8rem;
    color: var(--primary);
    text-align: center;
    padding: 1.5rem 0;
    font-weight: 700;
    margin-bottom: 0.5rem;
    background: linear-gradient(135deg, var(--primary), var(--secondary));
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}

.section-header {
    font-size: 1.6rem;
    color: var(--primary);
    border-left: 5px solid var(--secondary);
    padding-left: 1rem;
    margin: 2rem 0 1.5rem 0;
    font-weight: 600;
}

.panel {
    background-color: var(--card-bg);
    border-radius: var(--border-radius);
    padding: 1.8rem;
    box-shadow: var(--shadow);
    margin-bottom: 1.5rem;
    border: 1px solid #e0e0e0;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}

.panel:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 20px rgba(0, 0, 0, 0.12);
}

.metric-card {
    background: linear-gradient(135deg, var(--card-bg), #f7f9fc);
    border-radius: var(--border-radius);
    padding: 1.5rem;
    box-shadow: var(--shadow);
    text-align: center;
    border: 1px solid #e8e8e8;
    height: 100%;
}

.metric-value {
    font-size: 2rem;
    font-weight: 700;
    color: var(--primary);
    margin: 0.5rem 0;
}

.metric-label {
    font-size: 0.9rem;
    color: #666;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.result-table {
    background-color: var(--card-bg);
    border-radius: var(--border-radius);
    padding: 1.5rem;
    box-shadow: var(--shadow);
    margin: 1.5rem 0;
}

.optimization-formulation {
    background-color: #f8f9fa;
    border-left: 4px solid var(--primary);
    padding: 1.5rem;
    border-radius: 8px;
    margin: 1.5rem 0;
    font-family: 'Courier New', monospace;
    font-size: 0.9rem;
    line-height: 1.6;
}

.optimization-section {
    background-color: #fff;
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 1.5rem;
    margin: 1.5rem 0;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.stButton>button {
    background: linear-gradient(135deg, var(--primary), var(--secondary));
    color: white;
    border: none;
    padding: 0.8rem 2rem;
    border-radius: 50px;
    font-weight: 600;
    transition: all 0.3s ease;
    box-shadow: 0 4px 6px rgba(50, 50, 93, 0.11), 0 1px 3px rgba(0, 0, 0, 0.08);
    width: 100%;
    margin-top: 1rem;
}

.stButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0 7px 14px rgba(50, 50, 93, 0.1), 0 3px 6px rgba(0, 0, 0, 0.08);
    background: linear-gradient(135deg, var(--secondary), var(--primary));
}

.criteria-input {
    padding: 0.8rem;
    border-radius: 8px;
    border: 1px solid #e0e0e0;
    margin-bottom: 0.8rem;
    transition: border 0.3s ease;
    width: 100%;
}

.criteria-input:focus {
    border-color: var(--primary);
    outline: none;
    box-shadow: 0 0 0 2px rgba(31, 119, 180, 0.2);
}

.instruction-box {
    background-color: #f0f7ff;
    border-left: 4px solid var(--primary);
    padding: 1rem 1.5rem;
    border-radius: 4px;
    margin: 1rem 0;
    font-size: 0.95rem;
}

.success-box {
    background-color: #f0fff4;
    border-left: 4px solid var(--secondary);
    padding: 1rem 1.5rem;
    border-radius: 4px;
    margin: 1rem 0;
}

.warning-box {
    background-color: #fffaf0;
    border-left: 4px solid #ffb347;
    padding: 1rem 1.5rem;
    border-radius: 4px;
    margin: 1rem 0;
}

.error-box {
    background-color: #fff5f5;
    border-left: 4px solid #ff6b6b;
    padding: 1rem 1.5rem;
    border-radius: 4px;
    margin: 1rem 0;
}

.comparison-scale {
    background-color: #f8f9fa;
    padding: 1rem;
    border-radius: 8px;
    margin: 1rem 0;
    font-size: 0.9rem;
}

.scale-item {
    display: flex;
    justify-content: space-between;
    padding: 0.3rem 0;
    border-bottom: 1px dashed #e0e0e0;
}

.scale-item:last-child {
    border-bottom: none;
}

.footer {
    text-align: center;
    margin-top: 3rem;
    padding: 1.5rem;
    color: #666;
    font-size: 0.9rem;
    border-top: 1px solid #eaeaea;
}

.logo-container {
    text-align: center;
    margin-bottom: 1.5rem;
}

.logo {
    font-size: 3rem;
    margin-bottom: 0.5rem;
}

.stSlider .st-emotion-cache-1adqxcy {
    background: linear-gradient(to right, var(--primary), var(--secondary));
}

.stSelectbox div div div {
    color: var(--text);
}

input[type="number"] {
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 0.5rem;
}

</style>
""", unsafe_allow_html=True)

# ==================== OPA MODEL FUNCTIONS ====================

# Linguistic terms to TrFN for OPA
ling_to_trfn_opa = {
    'ELI': (1.0, 1.5, 2.5, 3.0),
    'VLI': (2.0, 2.5, 3.5, 4.0),
    'LI': (3.0, 3.5, 4.5, 5.0),
    'MI': (4.0, 4.5, 5.5, 6.0),
    'HI': (5.0, 5.5, 6.5, 7.0),
    'VHI': (6.0, 6.5, 7.5, 8.0),
    'EHI': (7.0, 7.5, 8.5, 9.0),
}

linguistic_options_opa = list(ling_to_trfn_opa.keys())

def trig_geom_component(values, weights):
    n = len(values)
    s = sum(values)
    if s == 0:
        return 0
    prod = 1
    for v, w in zip(values, weights):
        f = v / s
        prod *= np.sin(np.pi * f / 2) ** w
    arc = np.arcsin(prod)
    return s * (2 / np.pi) * arc

def aggregate_ftwg(trfn_list, weights):
    ls = [tf[0] for tf in trfn_list]
    ms = [tf[1] for tf in trfn_list]
    us = [tf[2] for tf in trfn_list]
    ws = [tf[3] for tf in trfn_list]
    
    l = trig_geom_component(ls, weights)
    m = trig_geom_component(ms, weights)
    u = trig_geom_component(us, weights)
    ww = trig_geom_component(ws, weights)
    
    return (l, m, u, ww)

def defuzz(trf):
    return (2 * trf[0] + 7 * trf[1] + 7 * trf[2] + 2 * trf[3]) / 18

def solve_fuzzy_opa(coeff_list, n):
    prob = LpProblem("Trapezoidal_Fuzzy_Trig_OPA", LpMaximize)
    
    # Define decision variables for the weights (w)
    w_l = [LpVariable(f"w_l_{i}", lowBound=0) for i in range(n)]
    w_m = [LpVariable(f"w_m_{i}", lowBound=0) for i in range(n)]
    w_u = [LpVariable(f"w_u_{i}", lowBound=0) for i in range(n)]
    w_w = [LpVariable(f"w_w_{i}", lowBound=0) for i in range(n)]
    
    # Define auxiliary variables for the objective function (Psi)
    Psi_l = LpVariable("Psi_l", lowBound=0)
    Psi_m = LpVariable("Psi_m", lowBound=0)
    Psi_u = LpVariable("Psi_u", lowBound=0)
    Psi_w = LpVariable("Psi_w", lowBound=0)
    
    # Objective Function: Maximize the defuzzified value of Psi
    prob += (2 * Psi_l + 7 * Psi_m + 7 * Psi_u + 2 * Psi_w) / 18
    
    # Constraints
    for i in range(n):
        prob += w_l[i] <= w_m[i]
        prob += w_m[i] <= w_u[i]
        prob += w_u[i] <= w_w[i]
    
    # Normalization constraints
    prob += lpSum(w_l) == 0.8
    prob += lpSum(w_m) == 0.9
    prob += lpSum(w_u) == 1.1
    prob += lpSum(w_w) == 1.2
    
    # Core OPA constraints comparing adjacent ranked criteria
    for a in range(n - 1):
        prob += coeff_list[a][0] * (w_l[a] - w_w[a + 1]) >= Psi_l
        prob += coeff_list[a][1] * (w_m[a] - w_u[a + 1]) >= Psi_m
        prob += coeff_list[a][2] * (w_u[a] - w_m[a + 1]) >= Psi_u
        prob += coeff_list[a][3] * (w_w[a] - w_l[a + 1]) >= Psi_w
    
    # OPA constraints for the last-ranked criterion
    prob += coeff_list[n - 1][0] * w_l[n - 1] >= Psi_l
    prob += coeff_list[n - 1][1] * w_m[n - 1] >= Psi_m
    prob += coeff_list[n - 1][2] * w_u[n - 1] >= Psi_u
    prob += coeff_list[n - 1][3] * w_w[n - 1] >= Psi_w
    
    # Solve the problem
    status = prob.solve()
    
    if status != 1:
        st.error("Optimization failed. The problem may be infeasible. Please check your inputs.")
        return None, None
    
    # Extract and clean results
    weights = []
    for i in range(n):
        wl = max(0, value(w_l[i]))
        wm = max(0, value(w_m[i]))
        wu = max(0, value(w_u[i]))
        ww = max(0, value(w_w[i]))
        weights.append((wl, wm, wu, ww))
    
    psi = (
        max(0, value(Psi_l)),
        max(0, value(Psi_m)),
        max(0, value(Psi_u)),
        max(0, value(Psi_w))
    )
    
    return weights, psi

def create_opa_word_document(criteria, theta, defuzz_values, coeff, ranked_criteria, weights, psi, num_experts, expert_weights):
    doc = Document()
    title = doc.add_heading('Trigonometric Trapezoidal Fuzzy OPA Results - Multiple Experts', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f'This document contains the results of the Fuzzy Trigonometric OPA analysis for {num_experts} experts.')
    doc.add_paragraph('')
    
    # Expert Weights
    doc.add_heading('Expert Weights', level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Expert'
    hdr_cells[1].text = 'Weight'
    
    for e in range(num_experts):
        row_cells = table.add_row().cells
        row_cells[0].text = f"Expert {e+1}"
        row_cells[1].text = f"{expert_weights[e]:.4f}"
    
    doc.add_paragraph('')
    
    # Aggregated Theta
    doc.add_heading('Aggregated Fuzzy Importance (Theta)', level=1)
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Criterion'
    hdr_cells[1].text = 'l'
    hdr_cells[2].text = 'm'
    hdr_cells[3].text = 'u'
    hdr_cells[4].text = 'w'
    hdr_cells[5].text = 'Defuzzified'
    
    for i, crit in enumerate(criteria):
        row_cells = table.add_row().cells
        row_cells[0].text = crit
        row_cells[1].text = f"{theta[i][0]:.4f}"
        row_cells[2].text = f"{theta[i][1]:.4f}"
        row_cells[3].text = f"{theta[i][2]:.4f}"
        row_cells[4].text = f"{theta[i][3]:.4f}"
        row_cells[5].text = f"{defuzz_values[i]:.4f}"
    
    doc.add_paragraph('')
    
    # Coefficients
    doc.add_heading('Coefficients for Fuzzy OPA', level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Criterion'
    hdr_cells[1].text = 'Coeff l'
    hdr_cells[2].text = 'Coeff m'
    hdr_cells[3].text = 'Coeff u'
    hdr_cells[4].text = 'Coeff w'
    
    for i, crit in enumerate(criteria):
        row_cells = table.add_row().cells
        row_cells[0].text = crit
        row_cells[1].text = f"{coeff[i][0]:.4f}"
        row_cells[2].text = f"{coeff[i][1]:.4f}"
        row_cells[3].text = f"{coeff[i][2]:.4f}"
        row_cells[4].text = f"{coeff[i][3]:.4f}"
    
    doc.add_paragraph('')
    
    # Ranked Criteria
    doc.add_heading('Ranked Criteria', level=1)
    table = doc.add_table(rows=1, cols=7)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Rank'
    hdr_cells[1].text = 'Criterion'
    hdr_cells[2].text = 'Weight l'
    hdr_cells[3].text = 'Weight m'
    hdr_cells[4].text = 'Weight u'
    hdr_cells[5].text = 'Weight w'
    hdr_cells[6].text = 'Defuzzified'
    
    for rank, (crit_idx, w) in enumerate(ranked_criteria):
        row_cells = table.add_row().cells
        row_cells[0].text = str(rank + 1)
        row_cells[1].text = criteria[crit_idx]
        row_cells[2].text = f"{w[0]:.4f}"
        row_cells[3].text = f"{w[1]:.4f}"
        row_cells[4].text = f"{w[2]:.4f}"
        row_cells[5].text = f"{w[3]:.4f}"
        row_cells[6].text = f"{defuzz(w):.4f}"
    
    doc.add_paragraph('')
    
    # Psi
    doc.add_heading('Psi Value', level=1)
    doc.add_paragraph(f'Psi: ({psi[0]:.4f}, {psi[1]:.4f}, {psi[2]:.4f}, {psi[3]:.4f})')
    doc.add_paragraph(f'Defuzzified Psi: {defuzz(psi):.4f}')
    
    doc_bytes = io.BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    return doc_bytes

def display_optimization_formulation(coeff_list, n, criteria_names):
    """Display the optimization problem formulation"""
    st.markdown('<div class="optimization-section">', unsafe_allow_html=True)
    st.markdown('<h3 class="section-header">TTrFS-OPA Optimization Problem Formulation</h3>', unsafe_allow_html=True)
    
    # Objective Function
    st.markdown("""
    <div class="optimization-formulation">
    <strong>Objective Function:</strong><br>
    Maximize: (2·Ψ_l + 7·Ψ_m + 7·Ψ_u + 2·Ψ_w) / 18
    </div>
    """, unsafe_allow_html=True)
    
    # Decision Variables
    st.markdown("""
    <div class="optimization-formulation">
    <strong>Decision Variables:</strong><br>
    w_lᵢ, w_mᵢ, w_uᵢ, w_wᵢ ≥ 0 for i = 1,...,n<br>
    Ψ_l, Ψ_m, Ψ_u, Ψ_w ≥ 0
    </div>
    """, unsafe_allow_html=True)
    
    # Constraints
    st.markdown("""
    <div class="optimization-formulation">
    <strong>Constraints:</strong><br>
    1. Fuzzy weight ordering:<br>
    &nbsp;&nbsp;w_lᵢ ≤ w_mᵢ ≤ w_uᵢ ≤ w_wᵢ for all i<br><br>
    
    2. Normalization constraints:<br>
    &nbsp;&nbsp;∑w_lᵢ = 0.8<br>
    &nbsp;&nbsp;∑w_mᵢ = 0.9<br>
    &nbsp;&nbsp;∑w_uᵢ = 1.1<br>
    &nbsp;&nbsp;∑w_wᵢ = 1.2<br><br>
    
    3. Adjacent criteria constraints (for a = 1 to n-1):<br>
    &nbsp;&nbsp;θₗᵃ·(w_lᵃ - w_wᵃ⁺¹) ≥ Ψ_l<br>
    &nbsp;&nbsp;θₘᵃ·(w_mᵃ - w_uᵃ⁺¹) ≥ Ψ_m<br>
    &nbsp;&nbsp;θᵤᵃ·(w_uᵃ - w_mᵃ⁺¹) ≥ Ψ_u<br>
    &nbsp;&nbsp;θ_wᵃ·(w_wᵃ - w_lᵃ⁺¹) ≥ Ψ_w<br><br>
    
    4. Last criterion constraint:<br>
    &nbsp;&nbsp;θₗⁿ·w_lⁿ ≥ Ψ_l<br>
    &nbsp;&nbsp;θₘⁿ·w_mⁿ ≥ Ψ_m<br>
    &nbsp;&nbsp;θᵤⁿ·w_uⁿ ≥ Ψ_u<br>
    &nbsp;&nbsp;θ_wⁿ·w_wⁿ ≥ Ψ_w
    </div>
    """, unsafe_allow_html=True)
    
    # Display specific coefficients for this problem
    st.markdown("""
    <div class="optimization-formulation">
    <strong>Problem-specific Coefficients (θ):</strong>
    </div>
    """, unsafe_allow_html=True)
    
    coeff_df = pd.DataFrame({
        'Criterion': criteria_names,
        'θ_l': [f"{coeff[0]:.4f}" for coeff in coeff_list],
        'θ_m': [f"{coeff[1]:.4f}" for coeff in coeff_list],
        'θ_u': [f"{coeff[2]:.4f}" for coeff in coeff_list],
        'θ_w': [f"{coeff[3]:.4f}" for coeff in coeff_list]
    })
    st.dataframe(coeff_df, use_container_width=True, hide_index=True)
    
    # Mathematical notation explanation
    with st.expander("Mathematical Notation Explanation"):
        st.markdown("""
        **Notation:**
        - **w_lᵢ, w_mᵢ, w_uᵢ, w_wᵢ**: Trapezoidal fuzzy weight components for criterion i
        - **Ψ_l, Ψ_m, Ψ_u, Ψ_w**: Auxiliary variables representing satisfaction degrees
        - **θₗᵢ, θₘᵢ, θᵤᵢ, θ_wᵢ**: Coefficient components for criterion i
        - **n**: Number of criteria
        
        **Coefficient Calculation:**
        θₗᵢ = min_l / wᵢ, θₘᵢ = min_l / uᵢ, θᵤᵢ = min_l / mᵢ, θ_wᵢ = min_l / lᵢ
        where min_l is the minimum l-component across all fuzzy importance values
        """)
    
    st.markdown('</div>', unsafe_allow_html=True)

def opa_model():
    st.markdown('<div class="logo-container">', unsafe_allow_html=True)
    st.markdown('<div class="logo">⚖️</div>', unsafe_allow_html=True)
    st.markdown('<h1 class="main-header">Trigonometric Trapezoidal Fuzzy OPA Analysis with Multiple Experts</h1>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div style="text-align: center; margin-bottom: 2.5rem; color: #555;">
    <p style="font-size: 1.1rem;">This application implements the Trigonometric Trapezoidal Fuzzy Ordinal Priority Approach (OPA) for multi-criteria decision-making with multiple experts.
    </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state for OPA
    if 'opa_criteria' not in st.session_state:
        st.session_state.opa_criteria = [f"Criterion {i+1}" for i in range(5)]
    if 'opa_num_criteria' not in st.session_state:
        st.session_state.opa_num_criteria = 5
    if 'opa_num_experts' not in st.session_state:
        st.session_state.opa_num_experts = 2
    if 'opa_results_calculated' not in st.session_state:
        st.session_state.opa_results_calculated = False
    if 'opa_expert_data' not in st.session_state:
        st.session_state.opa_expert_data = {}
    if 'opa_expert_weights' not in st.session_state:
        st.session_state.opa_expert_weights = [1.0 / st.session_state.opa_num_experts] * st.session_state.opa_num_experts
    
    # Create two columns for layout
    left_col, right_col = st.columns([1, 1])
    
    with left_col:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<h2 class="section-header">Step 1: Define Decision Criteria and Experts</h2>', unsafe_allow_html=True)
        
        num_experts = st.number_input(
            "Number of experts",
            min_value=1,
            max_value=15,
            value=st.session_state.opa_num_experts,
            step=1,
            key="opa_num_experts_input",
            help="Select between 1 to 15 experts"
        )
        
        num_criteria = st.number_input(
            "Number of criteria",
            min_value=3,
            max_value=25,
            value=st.session_state.opa_num_criteria,
            step=1,
            key="opa_num_criteria_input",
            help="Select between 3 to 25 criteria"
        )
        
        if num_criteria != st.session_state.opa_num_criteria or num_experts != st.session_state.opa_num_experts:
            st.session_state.opa_num_criteria = num_criteria
            st.session_state.opa_num_experts = num_experts
            if len(st.session_state.opa_criteria) < num_criteria:
                for i in range(len(st.session_state.opa_criteria), num_criteria):
                    st.session_state.opa_criteria.append(f"Criterion {i+1}")
            else:
                st.session_state.opa_criteria = st.session_state.opa_criteria[:num_criteria]
            st.session_state.opa_expert_data = {}
            st.session_state.opa_expert_weights = [1.0 / num_experts] * num_experts
            st.session_state.opa_results_calculated = False
        
        criteria = []
        for i in range(num_criteria):
            criterion = st.text_input(
                f"Criterion {i+1}",
                value=st.session_state.opa_criteria[i],
                key=f"opa_criterion_{i}",
                placeholder=f"Enter name for criterion {i+1}"
            )
            criteria.append(criterion)
        
        st.session_state.opa_criteria = criteria
        
        st.markdown('<h3>Expert Weights</h3>', unsafe_allow_html=True)
        st.markdown("""
        <div class="instruction-box">
        <strong>Important:</strong> The sum of all expert weights must equal exactly 1.00
        </div>
        """, unsafe_allow_html=True)
        
        expert_weights = []
        for e in range(num_experts):
            w = st.number_input(
                f"Weight for Expert {e+1}",
                min_value=0.0,
                max_value=1.0,
                value=st.session_state.opa_expert_weights[e],
                step=0.01,
                format="%.2f",
                key=f"opa_expert_w_{e}"
            )
            expert_weights.append(w)
        
        # Calculate and display sum of weights
        sum_w = sum(expert_weights)
        st.markdown(f"**Sum of expert weights: {sum_w:.2f}**")
        
        # Validate weights sum
        if abs(sum_w - 1.0) > 0.01:
            st.error(f"❌ Sum of expert weights must equal 1.00. Current sum: {sum_w:.2f}")
            st.session_state.opa_weights_valid = False
        else:
            # Normalize to ensure exact sum of 1.0
            if sum_w > 0:
                normalized_weights = [w / sum_w for w in expert_weights]
                st.success(f"✅ Weights are valid! Sum: {sum(normalized_weights):.2f}")
            else:
                normalized_weights = [1.0 / num_experts] * num_experts
                st.error("❌ Weights cannot be all zero!")
            st.session_state.opa_expert_weights = normalized_weights
            st.session_state.opa_weights_valid = True
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with right_col:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<h2 class="section-header">Expert Assessments</h2>', unsafe_allow_html=True)
        
        st.markdown("""
        <div class="instruction-box">
        <strong>Linguistic Terms:</strong><br>
        ELI (Extremely Low), VLI (Very Low), LI (Low), MI (Medium), HI (High), VHI (Very High), EHI (Extremely High)
        </div>
        """, unsafe_allow_html=True)
        
        # Individual expert tabs with data editor only
        tabs = st.tabs([f"Expert {e+1}" for e in range(num_experts)])
        
        for e, tab in enumerate(tabs):
            with tab:
                if f'expert_{e}' not in st.session_state.opa_expert_data:
                    st.session_state.opa_expert_data[f'expert_{e}'] = ['MI'] * num_criteria
                
                st.markdown(f'<h3>Expert {e+1} Input</h3>', unsafe_allow_html=True)
                
                # Data Editor for quick input
                st.markdown("**Edit ratings directly in the table below:**")
                
                # Create DataFrame for data editor
                df_data = {
                    'Criterion': criteria,
                    'Rating': st.session_state.opa_expert_data[f'expert_{e}']
                }
                df = pd.DataFrame(df_data)
                
                # Configure the data editor
                edited_df = st.data_editor(
                    df,
                    column_config={
                        "Rating": st.column_config.SelectboxColumn(
                            "Rating",
                            options=linguistic_options_opa,
                            required=True,
                            width="medium"
                        )
                    },
                    use_container_width=True,
                    key=f"opa_data_editor_expert_{e}"
                )
                
                # Update session state with edited data
                if not edited_df.equals(df):
                    st.session_state.opa_expert_data[f'expert_{e}'] = edited_df['Rating'].tolist()
                    st.success(f"Expert {e+1} ratings updated!")
                    st.rerun()
        
        # Calculate button with validation
        st.markdown('<div style="text-align: center; margin: 2.5rem 0;">', unsafe_allow_html=True)
        
        # Check if all conditions are met for calculation
        all_ratings_available = all(f'expert_{e}' in st.session_state.opa_expert_data for e in range(num_experts))
        weights_valid = getattr(st.session_state, 'opa_weights_valid', False)
        
        if not all_ratings_available:
            st.error("❌ Please provide ratings for all experts before calculating weights.")
            calculate_disabled = True
        elif not weights_valid:
            st.error("❌ Please fix expert weights (sum must equal 1.00) before calculating.")
            calculate_disabled = True
        else:
            calculate_disabled = False
        
        if st.button("Calculate Weights", key="opa_calculate_button", use_container_width=True, disabled=calculate_disabled):
            with st.spinner("Calculating weights..."):
                # Final validation before calculation
                if not all_ratings_available:
                    st.error("Please provide ratings for all experts before calculating weights.")
                elif not weights_valid:
                    st.error("Expert weights sum must equal 1.00. Please adjust the weights.")
                else:
                    # Aggregate theta using FTWG
                    theta = []
                    defuzz_values = []
                    for j in range(num_criteria):
                        trfn_list = [ling_to_trfn_opa[st.session_state.opa_expert_data[f'expert_{e}'][j]] for e in range(num_experts)]
                        aggregated_trf = aggregate_ftwg(trfn_list, st.session_state.opa_expert_weights)
                        theta.append(aggregated_trf)
                        defuzz_values.append(defuzz(aggregated_trf))
                    
                    # Compute min_l
                    min_l = min(t[0] for t in theta if t[0] > 0) if any(t[0] > 0 for t in theta) else 1.0
                    
                    # Compute coefficients
                    coeff = []
                    for t in theta:
                        if t[0] == 0 or t[1] == 0 or t[2] == 0 or t[3] == 0:
                            c = (0, 0, 0, 0)
                        else:
                            c = (min_l / t[3], min_l / t[2], min_l / t[1], min_l / t[0])
                        coeff.append(c)
                    
                    # Rank criteria based on defuzz(theta)
                    sorted_indices = np.argsort(defuzz_values)[::-1]
                    
                    # Sort coeff for OPA
                    coeff_sorted = [coeff[idx] for idx in sorted_indices]
                    
                    # Solve LP
                    weights_sorted, psi = solve_fuzzy_opa(coeff_sorted, num_criteria)
                    
                    if weights_sorted is not None:
                        # Map weights back to original order
                        weights = [None] * num_criteria
                        for rank, idx in enumerate(sorted_indices):
                            weights[idx] = weights_sorted[rank]
                        
                        # Ranked criteria with weights
                        ranked_criteria = [(sorted_indices[k], weights_sorted[k]) for k in range(num_criteria)]
                        
                        st.session_state.opa_theta = theta
                        st.session_state.opa_defuzz_values = defuzz_values
                        st.session_state.opa_coeff = coeff
                        st.session_state.opa_ranked_criteria = ranked_criteria
                        st.session_state.opa_weights = weights
                        st.session_state.opa_psi = psi
                        st.session_state.opa_results_calculated = True
                        st.success("Calculation completed successfully!")
        
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Results section
    if st.session_state.opa_results_calculated:
        st.markdown('<h2 class="section-header">Results</h2>', unsafe_allow_html=True)
        
        # Expert Weights
        st.markdown('<div class="result-table">', unsafe_allow_html=True)
        st.subheader("
