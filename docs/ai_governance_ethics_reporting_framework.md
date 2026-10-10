# Institutional Framework: Ethics, Transparency, Reporting & Audit Logging for AI in Financial Modelling

**Document Status**: Production Enterprise Standard  
**Governance Horizon**: 2026–2028 Statutory Compliance  
**Regulatory Baselines**: Federal Reserve SR 11-7, EU AI Act (Regulation 2024/1689), DORA (Regulation 2022/2554), GDPR (Article 22), MiFID II, SEC Predictive Data Analytics Proposal, FINRA Rules 3110/2010, ICAEW 20 Principles, FAST Standard.  
**Target Roles**: Chief Financial Officers (CFOs), Chief Risk Officers (CROs), Financial Controllers, Model Risk Managers, AI Engineering Leads, Compliance Committees.

---

## 1. Executive Summary & Core Governance Principles

As artificial intelligence transitions from conversational search assistants into core quantitative financial modeling, enterprise risk management requires a shift from informal prompting to a **governed, auditable architecture**. Passing unvalidated spreadsheets or financial feeds directly to Large Language Models (LLMs) creates severe systemic risks, including non-deterministic mathematical hallucinations, unrecorded baseline changes, and regulatory non-compliance.

This framework establishes the mandatory technical, ethical, operational, and audit standards for deploying AI across corporate financial planning and analysis (FP&A), capital budgeting, valuation, credit underwriting, and financial statement modeling.

### Core Architectural Mandates
1. **The Principle of Decoupled Computation**: *"Models reason, software computes, agents orchestrate, and systems verify."* Probabilistic LLMs must never function as primary financial calculators. All quantitative formulas, time-series aggregations, and accounting statements must be executed by deterministic code microservices (Python/SQL/Excel formula trees), restricting the LLM strictly to contextual reasoning, variance narrative synthesis, and structured proposal drafting.
2. **The Primary FP&A Governance Rule**: *"AI surfaces the draft; the human analyst owns the decision."* No AI system may autonomously commit updates to financial systems of record (General Ledger, ERP, Treasury platforms). All recommendations must pause at gated Human-in-the-Loop (HITL) checkpoints for explicit, authenticated review.
3. **Immutability of Audit Trails**: Every prompt, retrieved context passage, tool parameters, model reasoning step, human review action, and backend transaction must be captured in an append-only, cryptographically sealed audit log retained for regulatory inspection.

---

## 2. Ethical AI Principles in Financial Modelling

Financial decision-making directly impacts capital allocation, corporate solvency, employee compensation, shareholder value, and individual credit accessibility. Ethical AI governance prevents algorithmic bias, unfair discrimination, and strategic misalignment.

### 2.1 Algorithmic Fairness & Subgroup Bias Elimination
* **Fair Lending & Credit Alignment**: Models influencing creditworthiness, insurance underwriting, or capital allocation must undergo rigorous subgroup bias testing to ensure compliance with the Equal Credit Opportunity Act (ECOA) and EU AI Act Article 10.
* **Disparate Impact Thresholds**: Systems must achieve a **Disparate Impact Ratio $\ge 0.80$** across protected demographic categories and pass **Equalized Odds Ratio tests ($\ge 0.85$)** before release.
* **Prohibition of Proxy Variables**: Feature Selection Pipelines must scan and exclude proxy variables (e.g., geographic zip codes, device hardware metadata, network access times) that correlate with protected classes or perpetuate historical systemic biases.

### 2.2 Eliminating Strategic Misalignment
* **Misalignment Risk Definition**: Misalignment occurs when an AI model successfully optimizes a narrow, localized objective function (e.g., maximizing short-term cash balances by freezing capital projects) that directly conflicts with executive leadership's strategic priorities (e.g., long-term R&D investment).
* **Multi-Scenario Sensitivity Checks**: All AI-driven scenario generators must evaluate trade-offs across range-bound probability distributions ($	ext{P10}$, $	ext{P50}$, $	ext{P90}$) and Cash Flow at Risk ($	ext{CFaR}$) metrics rather than presenting single-point deterministic outputs that disguise tail risks.

---

## 3. Technical Transparency & Explainability Architecture (XAI)

Deep learning networks and complex ensemble models operate as "black boxes" containing millions of uninterpretable weights. To satisfy Federal Reserve SR 11-7 Pillar 1 (Conceptual Soundness) and EU AI Act Articles 13 & 14, institutions must deploy post-hoc Explainable AI (XAI) frameworks.

### 3.1 Feature Attribution via SHAP & LIME
```
                        SR 11-7 & EU AI ACT XAI PIPELINE
                                       │
 ┌─────────────────────────────┼─────────────────────────────┐
 ▼                             ▼                             ▼
GLOBAL TRANSPARENCY           FEATURE DRIFT MONITORING      LOCAL EXPLANABILITY (LIME)
• SHAP Value Decomposition    • Population Stability Index  • Adverse Action Disclosures
• Global Factor Rankings      • 30-Day Rolling PSI Threshold• Top 3-5 Negative Factor Drivers
• Theoretical Alignment       • PSI ≥ 0.25 Forces Fallback  • Statutory "Right to Explanation"
```

1. **Global Factor Attributability (SHAP)**: Compute mean absolute **SHAP (SHapley Additive exPlanations)** values across training and validation sets. Model Risk Managers must verify that top drivers (e.g., Debt-to-Equity ratio, interest coverage) align with established financial theory rather than spurious correlations.
2. **Local Factor Attributability (LIME)**: For every individual automated credit denial, budget variance flag, or risk re-rating, **LIME (Local Interpretable Model-agnostic Explanations)** must generate the exact top 3–5 factor contributions driving the specific decision, satisfying GDPR Article 22 "Right to Explanation" requirements.

### 3.2 Feature Attribution Drift & PSI Threshold Guardrails
To detect concept drift and economic regime shifts post-deployment, quantitative teams must track the **Population Stability Index (PSI)** of production feature attributions against baseline training distributions:

$$	ext{PSI} = \sum \left( 	ext{Actual}\% - 	ext{Expected}\% ight) 	imes \ln\left( rac{	ext{Actual}\%}{	ext{Expected}\%}ight)$$

* **$	ext{PSI} < 0.10$**: **Stable**; normal operation.
* **$0.10 \le 	ext{PSI} < 0.25$**: **Moderate Shift**; triggers automated alert for Model Risk Manager review.
* **$	ext{PSI} \ge 0.25$**: **Severe Shift**; automatically forces model fallback to conservative, rule-based baseline logic.

### 3.3 Data Lineage & Context Compaction Guardrails
* **Context Lineage Preservation**: When long-horizon agent harnesses compress context windows via summarization or compaction, source metadata (document ID, section, page, General Ledger transaction ID) must be preserved in a queryable index outside the context window rather than converted into raw paraphrased prose.
* **Retrieval Augmented Generation (RAG) Grounding**: Qualitative variance narratives must draw exclusively from indexed, certified enterprise knowledge bases (e.g., General Ledger notes, SEC filings, contract terms). Any claim uncited to a retrieved passage is automatically flagged as ungrounded and rejected.

---

## 4. Human-in-the-Loop (HITL) & Operational Control Boundaries

Enterprise safety requires establishing strict operational envelopes that separate autonomous search from gated financial writes.

### 4.1 Operational Read vs. Write Boundaries
| Operational Tier | Execution Mode | Allowed Actions | Governance & Security |
| :--- | :--- | :--- | :--- |
| **Tier 1: Read & Analyze** | **Autonomous** | Pulling actuals, retrieving policy contracts, computing EVM ratios (CPI, SPI, EAC), running Monte Carlo simulations. | Authenticated via On-Behalf-Of (OBO) user tokens; read-only access. |
| **Tier 2: Recommend** | **Autonomous** | Generating forecast adjustments, drafting board pack narratives, identifying cost anomalies. | RAG grounding required; output presented as staged proposal. |
| **Tier 3: Write & Commit** | **Gated (HITL Sign-off)** | Updating General Ledger baselines, committing budget reserves, releasing funds, submitting regulatory filings. | System halts execution; requires explicit approval card click + reviewer notes. |

### 4.2 Interactive HITL Approval Gateway
When an agent recommends a financial adjustment (e.g., updating a project's Estimate at Completion):
1. The agent executes `propose_forecast_adjustment(...)` rather than calling an ERP write endpoint.
2. The workflow pauses stateful execution (via Celery/Redis or HTTP 202 async pause).
3. A structured **Adaptive Card** is delivered to the controller (via Microsoft Teams or financial portal), displaying:
   * Current baseline vs. proposed figure (with delta).
   * Underlying financial calculation breakdown.
   * RAG contract evidence & justification narrative.
   * Explicit **Approve**, **Reject**, and **Request Modification** action buttons.
4. Upon human sign-off, the backend executes the write operation and logs the complete transaction.

### 4.3 Identity Delegation Architecture
* **User Query Phase**: Conversational requests use **On-Behalf-Of (OBO)** OAuth 2.0 / Entra ID tokens so the AI agent strictly inherits the individual user's data access permissions.
* **System Execution Phase**: Approved backend writes execute under a **Dedicated Agent User ID** (e.g., `Agent_Finance_Execution_Service`). This ensures that ERP audit logs cleanly distinguish human-initiated manual edits from automated agent executions approved by a controller.

---

## 5. Cryptographic Audit Logging & Lineage Specification

To satisfy **EU AI Act Article 12 (Automatic Logging)**, **Article 26 (Deployer Log Retention)**, and **SR 11-7**, all AI interactions must produce tamper-evident, append-only audit records.

### 5.1 PostgreSQL DDL: Immutable Audit Log Schema (`schema_audit_log.sql`)
```sql
-- PostgreSQL DDL for Cryptographically Chained Financial Audit Trail
CREATE TABLE IF NOT EXISTS financial_proposal_audit_log (
    audit_id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    proposal_id UUID NOT NULL,
    project_id VARCHAR(64) NOT NULL,
    wbs_code VARCHAR(64) NOT NULL,
    
    -- Transaction Details
    action_type VARCHAR(32) NOT NULL, -- 'STAGED_PROPOSAL', 'HUMAN_APPROVAL', 'HUMAN_REJECTION', 'ERP_COMMIT'
    proposed_amount NUMERIC(18, 2) NOT NULL,
    currency VARCHAR(8) NOT NULL DEFAULT 'USD',
    reason_code VARCHAR(64) NOT NULL,
    
    -- Security & Identity Isolation
    requested_by_user_id VARCHAR(128) NOT NULL, -- On-Behalf-Of (OBO) Identity
    reviewed_by_approver_id VARCHAR(128),      -- Authenticated Approver Entra ID
    executed_by_service_id VARCHAR(128),       -- Dedicated Service ID
    
    -- Context & Evidence Snapshots
    evm_context_snapshot JSONB NOT NULL,       -- CPI, SPI, BAC metrics
    justification_narrative TEXT NOT NULL,      -- RAG evidence + reasoning
    approver_comments TEXT,                    -- Reviewer audit notes
    
    -- Timestamps & Cryptographic Hash Chaining
    created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
    previous_record_hash VARCHAR(64),          -- SHA-256 hash of prior log entry
    record_hash VARCHAR(64) NOT NULL           -- SHA-256 hash of current entry
);

-- Prevent Post-Hoc Alteration or Deletion of Audit Records
CREATE OR REPLACE FUNCTION prevent_audit_log_tampering()
RETURNS TRIGGER AS $$
BEGIN
    RAISE EXCEPTION 'CRITICAL SECURITY VIOLATION: Audit log records are immutable and cannot be updated or deleted.';
    RETURN NULL;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER enforce_audit_immutability
BEFORE UPDATE OR DELETE ON financial_proposal_audit_log
FOR EACH ROW EXECUTE FUNCTION prevent_audit_log_tampering();
```

### 5.2 Cryptographic Hash Chaining Algorithm
Each log record computes a SHA-256 hash incorporating the prior record's hash, preventing post-hoc log alteration or deletion:

$$	ext{Record Hash}_t = 	ext{SHA256}\Big(	ext{Record Hash}_{t-1} + 	ext{proposal\_id} + 	ext{action\_type} + 	ext{amount} + 	ext{approver\_id} + 	ext{timestamp}\Big)$$

If an auditor re-computes the hash chain and detects a mismatch, the log is flagged as tampered.

---

## 6. Model Evaluation, Daily Regression Testing & Quality Gates

To prevent prompt drift, vector indexing decay, and hallucination propagation, institutions must deploy **Standing Daily Evaluation Gates** (The Morgan Stanley Pattern).

```
                      STANDING DAILY EVALUATION HARNESS
                                      │
 ┌────────────────────────────────────┼────────────────────────────────────┐
 ▼                                    ▼                                    ▼
1. ACCURACY & PRECISION             2. LOGICALITY & REASONING             3. STORYTELLING & STRUCTURE
• Factual & Numeric Accuracy        • Coherence of Financial Deductions   • Structural Formatting
• Error Metrics: MAE, RMSE, MAPE    • Parameter-Passing Precision         • Board-Pack Compliance
• Zero Math Hallucinations          • Correct Tool Call Selection         • Actionable Action Titles
```

### 6.1 Core Evaluation Metrics
1. **Numeric & Factual Accuracy**:
   * **Mean Absolute Error (MAE)**: Measures average absolute prediction miss in native currency units.
   * **Root Mean Squared Error (RMSE)**: Penalizes large outliers heavily.
   * **Mean Absolute Percentage Error (MAPE)**: Relative percentage error across business units.
   * **Mean Absolute Scaled Error (MASE)**: Evaluates forecast performance against a seasonal naive baseline ($< 1.0$ indicates AI adds value).
2. **Probabilistic Calibration Metrics**:
   * **Continuous Ranked Probability Score (CRPS)** & **Pinball Quantile Loss**: Measures accuracy and sharpness of range-bound probabilistic forecasts ($	ext{P10}/	ext{P50}/	ext{P90}$).
3. **Qualitative & Narrative Scoring (0–10 Scale)**:
   * **Accuracy Score**: Numeric precision and RAG passage citation accuracy.
   * **Logicality Score**: Coherence of financial cause-and-effect deductions.
   * **Storytelling Score**: Structural clarity and alignment with sell-side/CFO reporting standards.

### 6.2 Standing CI/CD Release Gates
Before any model update, system prompt change, or vector index refresh is deployed to production, it must pass an automated CI/CD evaluation pipeline:
* **Accuracy Gate**: 100% pass rate on deterministic math regression tests.
* **Hallucination Gate**: Zero uncited factual claims in RAG narrative tests.
* **Safety Gate**: 100% rejection of adversarial prompt injection attacks.

---

## 7. Reporting & In-Workbook Governance (Excel Integration)

When financial models produce Microsoft Excel deliverables (`.xlsx`), the workbook structure must embed automated integrity checks aligned with **ICAEW Principle #19** and **SSRB Standard 11**.

```
┌──────────────────────────────────────────────────────────────────────────────────┐
│                   FINANCIAL MODEL INTEGRITY & AUTOMATED ERROR CHECKS             │
├──────────────────────────────────────────────────────────────────────────────────┤
│ ⚠️ / ✅ MASTER STATUS: ALL AUDIT CHECKS PASSED  (Frozen Banner Row 2)             │
├──────────┬──────────────────────────────────────────────┬─────────────┬──────────┤
│ Check ID │ Assertion Description                        │ Error Count │ Status   │
├──────────┼──────────────────────────────────────────────┼─────────────┼──────────┤
│ CHK-01   │ Balance Sheet Equality (Assets - L&E = 0)   │      0      │ PASS ✅  │
│ CHK-02   │ Cash Flow Reconciliation (CF Cash = BS Cash) │      0      │ PASS ✅  │
│ CHK-03   │ Formula Error Scan (#REF!, #DIV/0!, #VALUE!) │      0      │ PASS ✅  │
│ CHK-04   │ Retained Earnings Roll-Forward Equality      │      0      │ PASS ✅  │
│ CHK-05   │ Unwrapped Division Guardrail (#DIV/0! Trap)  │      0      │ PASS ✅  │
└──────────┴──────────────────────────────────────────────┴─────────────┴──────────┘
```

### Mandatory In-Workbook Controls
1. **Dedicated `Model_Checks` Tab**: Consolidates all automated model checks into a single matrix.
2. **Master Freeze-Pane Banner**: Row 2 of every worksheet displays a master calculation formula referencing the `Model_Checks` tab:
   ```excel
   =IF(SUM(Model_Checks!C5:C10)>0, "⚠️ MODEL STATUS: ERRORS DETECTED", "✅ MODEL STATUS: ALL AUDIT CHECKS PASSED")
   ```
3. **Tolerance-Based Checks**: Apply rounding tolerances (`ABS(Calculated - Control) < 0.01`) to prevent floating-point representation errors from triggering false positives.
4. **Conditional Formatting**: Triggered errors automatically format in bold red text (`#FCE4D6` fill), while passing checks display in bold green text (`#E2EFDA` fill).

---

## 8. Multi-Regulatory Compliance Cross-Walk

This matrix details how the technical controls in this framework map directly to major global financial regulations:

| Regulatory Framework | Specific Article / Rule | Mandatory Technical Requirement | Framework Implementation Section |
| :--- | :--- | :--- | :--- |
| **EU AI Act** | **Annex III (Point 5b/c)** | High-risk classification for credit/underwriting AI. | Section 1 (Risk Tiering) |
| **EU AI Act** | **Article 10** | Data governance, representativeness, bias testing. | Section 2.1 (Subgroup Bias) |
| **EU AI Act** | **Article 12 & 26** | Automatic decision logging & minimum 6-month log retention. | Section 5 (Audit Logging) |
| **EU AI Act** | **Article 14** | Operable human oversight interfaces & override authority. | Section 4 (HITL Gateway) |
| **Federal Reserve** | **SR 11-7 (Pillars 1–3)** | Conceptual soundness, ongoing monitoring, outcomes analysis. | Sections 3, 5, 6 |
| **GDPR** | **Article 22** | "Right to Explanation" for automated individual decisions. | Section 3.1 (LIME Attribution) |
| **DORA** | **Article 5–15** | ICT risk management & critical third-party service provider logging. | Section 5 (PostgreSQL DDL) |
| **MiFID II** | **Article 16(7)** | 5-year immutable record-keeping of investment communications. | Section 5 (Cryptographic Chaining) |
| **SEC** | **Predictive Data Analytics** | Elimination of algorithmic revenue-prioritization conflicts. | Section 2.2 (Misalignment) |
| **FINRA** | **Rules 3110 / 2010** | Supervisory control systems & firm legal accountability. | Section 4.3 (Dual Identity) |

---

## 9. Conclusion & Actionable Implementation Roadmap

1. **Step 1: Inventory & Classify**: Catalog all financial AI touchpoints and classify them into Low-Risk vs. High-Risk use cases.
2. **Step 2: Deploy XAI & RAG Layer**: Mandate SHAP/LIME feature attribution and restrict qualitative commentary to RAG-indexed corporate repositories.
3. **Step 3: Enforce HITL Approval Gates**: Insert structured proposal checkpoints (Adaptive Cards) before any financial update reaches systems of record.
4. **Step 4: Instrument Cryptographic Audit Logging**: Deploy append-only, SHA-256 chained PostgreSQL schemas to log all decision events.
5. **Step 5: Activate Daily Evaluation Harnesses**: Run automated daily regression test suites to catch prompt drift and mathematical hallucinations prior to executive reporting.
