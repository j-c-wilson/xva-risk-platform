
# Multi-Asset Valuation, xVA & Risk Analytics Platform

  

>  **Technicals:** Microsoft Excel, Custom Formulas, What-If Analysis, Conditional Formatting, Heatmaps, Charting & Dashboards, VBA & Automation, VBA Macros, Monte Carlo Simulations, Stress Testing & Scenario Analysis.

  

This platform is an institutional-grade front-office risk engine designed to perform multi-asset portfolio valuation, counterparty credit risk assessment, xVA capital budgeting, and stress testing. The system combines an in-memory 10,000-trial Monte Carlo simulation engine written in vectorized VBA with dynamic Excel dashboard controls and closed-form analytic models.

![alt text](https://github.com/j-c-wilson/xva-risk-platform/blob/main/Dashboard%20View.png "Dashboard View")
  
  

## Features

  

### Dashboard Master Controls & Toggle Switches

  

* **Multi-Position Toggle:**

* Toggles between single-asset exposure and multi-asset portfolio netting.

* When enabled, models structural portfolio netting benefits under an ISDA Master Agreement.

* **Collateral Haircut Toggle:**

* Toggles asset liquidation haircuts on or off to isolate pure credit/market risk from liquidity-adjusted collateral calls.

* **Master Volatility Multiplier %:**

* Global scalar that scales market volatility across all asset positions simultaneously to simulate macro market stress.

  

### Asset & Portfolio Input Parameters

  

* **Asset Class Selection:** Supports cash, corporate bonds, commodities, equities, and cryptocurrencies, featuring specific asset class liquidation haircuts (e.g., 0.00% for Cash, 8.00% for Corporate Bonds/Crypto).

* **Notional Position Sizing & Volatility:** Configurable trade notionals and independent annual volatility controls per position, dynamically scaled to holding horizons (e.g., 10-day holding period).

* **Correlation Coefficient (ρ):** Inter-asset correlation parameter controlling joint price movements and portfolio diversification.

* **Target Confidence Level (α):** User-adjustable quantile threshold driving VaR, Expected Shortfall, and Initial Margin calculations.

  

### 10,000-Trial Monte Carlo Engine

  

* **Correlated Path Generation:** Employs **Cholesky Decomposition** to generate dependent standard normal variates when multiple positions are enabled:

 $$ X_1 = Z_1, \quad X_2 = \rho \cdot X_1 + \sqrt{1 - \rho^2} \cdot Z_2 $$ 
 

* **VBA Macro (`MasterMonteCarloSim`):** Custom macro that processes all 10,000 paths inside 64-bit RAM arrays and writes to the worksheet in a single range assignment, drastically reducing simulation runtime.

  

### Output Metrics & Full xVA Decomposition

  

* **Target Confidence Value at Risk ($\text{VaR}_\alpha$):** Horizon tail loss at specified confidence.

* **Expected Shortfall ($\text{ES}_\alpha$ / Conditional VaR):** Average loss strictly beyond the VaR cutoff.

* **Gross Required Collateral Buffer Today:** Total cash/collateral call required after applying asset haircuts.

* **Net xVA Portfolio Change:** Total valuation adjustment applied to risk-free MtM.

* **Complete xVA Breakdown:**

	* **CVA:** Counterparty default risk pricing via Expected Positive Exposure ($\text{EPE}$).

	* **DVA:** Own credit risk pricing via Expected Negative Exposure ($\text{ENE}$).

	* **FVA:** Cost of uncollateralized funding spreads.

	* **MVA:** Cost of funding posted Initial Margin ($\text{IM}$) over trade lifetime.

	* **KVA:** Cost of holding regulatory capital ($K_{\text{reg}}$) at the corporate hurdle rate.

  

### Dashboard Visualizations & Stress Testing Heatmap

  

* **Monte Carlo P&L Distribution Chart:** Histogram visualizing profit and loss outcomes across 10,000 trials.

* **xVA Capital Breakdown Chart:** Donut chart breaking down capital allocation across CVA, DVA, FVA, MVA, and KVA.

* **Horizon Exposure vs. Time Horizon Graph:** Line chart plotting collateral buffer growth across horizons ($0$ to $365$ days).

* **Initial Margin (IM) Sensitivity Heatmap:** Dynamic 2D Data Table matrix evaluating initial margin requirements across correlation shifts ($\rho$) and volatility multipliers. Relative conditional formatting maps low risk (green), baseline (yellow), and severe stress (red).

  



  

## xVA Valuation Adjustments

  

The platform incorporates a full valuation adjustment framework to price counterparty, funding, margin, and regulatory capital costs into the portfolio:


| xVA Component | Primary Risk Driver | Formula / Logic Breakdown |
| :--- | :--- | :--- |
| **CVA** (Credit Valuation Adj.) | Counterparty Default Risk | $(1 - R_c) \cdot \text{PD}_c \cdot \text{EPE}$ <br> • $R_c$: Recovery rate of counterparty <br> • $\text{PD}_c$: Default probability of counterparty <br> • $\text{EPE}$: Expected Positive Exposure (uncollateralized credit risk) |
| **DVA** (Debit Valuation Adj.) | Own Default Risk (Bilateral) | $(1 - R_{\text{own}}) \cdot \text{PD}_{\text{own}} \cdot \text{ENE}$ <br> • <i>R</i><sub>own</sub>: Own institution recovery rate <br> • <b>PD</b><sub>own</sub>: Own default probability <br> • <b>ENE</b>: Expected Negative Exposure (our debt to counterparty) |
| **FVA** (Funding Valuation Adj.) | Unsecured Borrowing/Lending Spread | $\text{Funding Spread} \cdot (\text{EPE} - \text{ENE})$ <br> • $\text{Funding Spread}$: Net cost to borrow/lend unsecured capital <br> • $(\text{EPE} - \text{ENE})$: Net uncollateralized portfolio exposure |
| **MVA** (Margin Valuation Adj.) | Cost of Posting Initial Margin ($\text{IM}$) | $\text{Cost of Funding IM} \cdot \sum \mathbb{E}[\text{IM}_t]$ <br> • $\text{Cost of Funding IM}$: Spread cost to hold/post regulatory margin <br> • $\sum \mathbb{E}[\text{IM}_t]$: Sum of expected Initial Margin requirements across trade lifetime |
| **KVA** (Capital Valuation Adj.) | Regulatory Capital Hurdle Rate | $\text{Hurdle Rate} \cdot \sum K_{\text{reg}}$ <br> • $\text{Hurdle Rate}$: Target return rate required by shareholders <br> • $\sum K_{\text{reg}}$: Cumulative regulatory capital required over trade lifetime |

 

  

## xVA Across Varying Confidence Levels

  

A core finding from this model is the structural interaction between CVA and MVA across varying confidence levels:

![alt text](https://github.com/j-c-wilson/xva-risk-platform/blob/main/Confidence%20vs%20xVa.png "Varying Confidence")

### Low Confidence

At lower confidence levels ($\alpha$), the required Initial Margin ($\text{IM}$) is relatively small. Because collateral coverage is low, Expected Positive Exposure ($\text{EPE}$) remains large. As a result, Counterparty Credit Risk ($CVA$) dominates the valuation adjustment, while the cost of carrying margin ($MVA$) is negligible:

  

$$\text{Low } \alpha \longrightarrow \text{Low IM} \longrightarrow \text{High EPE} \longrightarrow \mathbf{\text{CVA} > \text{MVA}}$$

  

### High Confidence

As target confidence increases, required $\text{IM}$ grows non-linearly, driven by the tail quantile function ($\text{NORM.S.INV}$). High collateral coverage absorbs tail risk, driving expected uncollateralized exposure near zero and virtually eliminating $CVA$. However, funding that massive collateral buffer creates high carrying costs. Thus, $MVA$ expands exponentially while $CVA$ drops to zero:

$$\text{High } \alpha \longrightarrow \text{High IM} \longrightarrow \text{EPE} \approx 0 \longrightarrow \mathbf{\text{MVA} \gg \text{CVA}}$$


  

## Institutional Applications

  

This platform mirrors the core quantitative functions of institutional trading desks and risk management units:

  

* **Pre-Trade Pricing:** Calculates exact marginal xVA adjustments prior to trade execution.

* **Margin Optimization:** Evaluates how asset correlation shifts alter initial margin requirements under ISDA rules.

* **Treasury Liquidity Buffer Sizing:** Enables corporate treasurers to run forward stress tests to prevent liquidity crunches under market shocks.
