# Antibiotic Resistance Analysis Tool

This tool analyzes antibiotic susceptibility data to determine:
- Effectiveness of individual antibiotics
- Effectiveness of antibiotic combinations
- Global coverage of combinations

## Excel File Requirements

### File Structure
Your Excel file must be structured as follows:

1. **First column**: Must be named "Microorganism" and contain microorganism names/IDs
2. **Subsequent columns**: Each column represents an antibiotic
3. **Data values**: 
   - `S` = Susceptible
   - `R` = Resistant
   - `N` = Not tested
   - `I` = Intermediate (will be automatically converted to `S`)

### Example Structure

| Microorganism | Ampicillin | Gentamicin | Ciprofloxacin | ... |
|---------------|------------|------------|---------------|-----|
| E. coli       | S          | R          | S             | ... |
| S. aureus     | R          | S          | N             | ... |
| ...           | ...        | ...        | ...           | ... |

### Important Notes
1. Only `S`, `R`, and `N` values are allowed (case insensitive)
2. Empty cells will be treated as `N` (Not tested)
3. The file must be in `.xlsx` format

### Supporting Files

This tool requires additional Excel files for specific functions:

* **Gram Stain Map (`gram_stain_map.xlsx`)**: A simple Excel file with two columns, `Microorganism` and `Gram_stain`. The script uses this to classify bacteria. If a microorganism is not found, the tool will ask for the classification and update the file for future use.
* **Exclusion List (Optional)**: A single-column Excel file containing a list of microorganism names to be excluded from the analysis (e.g., contaminants).

---

## How to Use

1.  Run the script (Google Colab is recommended).
2.  Upload your primary data file when prompted.
3.  Upload your Gram stain map file. The tool will automatically create a new one if it doesn't exist.
4.  Specify if you wish to **upload a list of microorganisms to exclude** from the analysis.
5.  Enter the **minimum testing threshold** (0-100): This filters out antibiotics tested on less than X% of the isolates.
6.  Enter the desired **confidence level** (e.g., 95 for 95% CI).

The tool will then perform the analysis and export the results to an Excel file.

---

## Analytical Metrics and Statistics

For each antibiotic and two-drug combination, the following metrics are calculated:

* **Susceptibility Percentage**: The proportion of susceptible isolates relative to the total number of tested isolates (S / [S+R]).
* **Global Effectiveness**: The proportion of susceptible isolates relative to the total number of all isolates in the dataset (S / total isolates).
* **Confidence Intervals**: Calculated using the beta distribution at the specified confidence level to quantify the precision of the estimates.

All metrics are calculated for three distinct data subsets: all isolates, Gram-positive isolates, and Gram-negative isolates.

---

## Output

The script generates a single `.xlsx` file with the following sheets:

* **Summary**: A comprehensive overview that includes:
    * The total number of excluded isolates.
    * The distribution of Gram-positive vs. Gram-negative isolates.
    * A list of the top 10 most common microorganisms.
    * Lists of the top 10 most resistant antibiotics for both Gram-positive and Gram-negative subsets.
* **Separate Sheets**: The full analytical results are provided on dedicated sheets for each data subset:
    * `All - Singles` and `All - Combs`
    * `Gram+ - Singles` and `Gram+ - Combs`
    * `Gram- - Singles` and `Gram- - Combs`

---

## Requirements

* Python 3.7+
* Required packages: `pandas`, `numpy`, `scipy`, `openpyxl`
* For Google Colab: no additional setup is needed.
