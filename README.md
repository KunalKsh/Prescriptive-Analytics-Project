# Prescriptive-Analytics-Project
This project uses Excel Solver to optimize profit by assigning dealers to insurance branches. It's a MILP that calculates the best match based on a profit matrix and constraints, including branch capacity and business-rule feasibility. The goal is to maximize total profit from all assignments.

### README: Excel-Based Assignment Problem Solver
The solution is broken down into a structured, multi-sheet Excel workbook.

### Project Files

* **`PRESCRIPTIVE ANALYTICS PROJECT.xlsx`**: The primary Excel file containing all data, formulas, and the Solver model.

### Workbook Structure

The workbook is organized into four key sheets to separate data from the model logic:

* **`Inputs`**: Contains global constants, such as the flat Cost Per Assignment (CPA).
* **`Dealers`**: A data table listing all dealers and their attributes (Premium, Minimum Payout, RTO, OEM, Fuel, Body Type).
* **`Branches`**: A data table detailing each branch, its attributes (PO%, Capacity), and business rules for feasibility.
* **`Feasibility`**: A pre-computed binary matrix (0s and 1s) acting as a mask. A `1` indicates a dealer-branch combination is valid based on all business and financial rules; a `0` means it's not. This simplifies the Solver model.
* **`Model`**: The core of the solver. It contains three matrices:
    * **Profit Matrix (`Mij`)**: Calculates the potential profit for every possible dealer-branch assignment using the formula `(Premium * PO%) - Dealer Min - CPA`.
    * **Decision Variables (`Xij`)**: A binary matrix where Solver places a `1` for an assigned pair and a `0` otherwise.
    * **Objective Function**: A single cell that uses `SUMPRODUCT` to calculate the total profit (`Mij * Xij`) and is set to be maximized by Solver.

### How It Works

1.  **Data Input**: All dealer and branch information is entered into their respective sheets.
2.  **Feasibility Check**: A `0/1` mask is manually or formulaically created in the `Feasibility` sheet to filter out invalid assignments based on the provided business rules (e.g., a branch not supporting a specific RTO or OEM).
3.  **Profit Calculation**: The `Model` sheet calculates the potential profit for all valid assignments.
4.  **Solver Execution**:
    * The objective is set to **maximize** the total profit cell.
    * The variable cells are set to the `Xij` matrix.
    * **Constraints** are added to ensure:
        * Each dealer is assigned to exactly one branch (`SUM(Xij)` per row = 1).
        * Each branch does not exceed its capacity (`SUM(Xij)` per column <= Capacity).
        * Assignments are only made for feasible pairs (`Xij <= Feasibilityij`).
        * The decision variables are binary (`Xij` = 0 or 1).
5.  **Result**: Solver finds the optimal combination of assignments that yields the highest total profit, populating the `Xij` matrix with the solution.
