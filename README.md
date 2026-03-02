Now I have everything I need. Here's the README:

---

# Excel VBA Automation — Financial Modeling & Simulation

Excel VBA automation applied to financial modeling, built as part of **FIN 645** at the University of Tampa.

> **Disclosure:** Tutorial structure and guided instructions provided as part of FIN 645 coursework. VBA code, custom functions, and simulation logic written independently.

---

## What's in This Repo

| File | Description |
|---|---|
| `2023_Excel_VBA_Tutorial_SOLUTION.xlsm` | Completed workbook with all macros, simulations, and functions |

---

## What Was Built

### Monte Carlo Simulation
Wrote a VBA subroutine that runs 1,000 iterations of a 30-year portfolio return simulation, generating random returns using a normal distribution and storing results programmatically in column K — replacing manual recalculation entirely.

### Custom User Defined Functions (UDFs)
Built a `FV_Volatility` function applying continuous compounding with volatility penalty:

```
FV = PV × e^((R - 0.5σ²) × T)
```

### Payment-Based Volatility Model
Macro-driven payment accumulation model using geometric return with volatility erosion, triggered via form controls.

### Solver Automation — Portfolio Optimization
Integrated VBA with Excel Solver to maximize the Sharpe Ratio by dynamically adjusting asset weights, replicating an institutional portfolio optimization workflow.

### VBA Fundamentals
For/Next loops, subroutines, cell referencing (Range, Cells, ActiveCell), button-triggered macros, and relative vs absolute references.

---

## Skills Demonstrated

`VBA` `Excel Automation` `Monte Carlo Simulation` `User Defined Functions` `Solver Integration` `Portfolio Optimization` `Continuous Compounding` `Loop Construction` `Financial Modeling`

---

## Source

Developed as part of **FIN 645** at the University of Tampa.

