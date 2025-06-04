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

## How to Use

1. Run the script (Google Colab recommended)
2. Upload your Excel file when prompted
3. Enter the minimum testing threshold (0-100):
   - This filters out antibiotics tested in less than X% of cases
   - Example: 50 = only include antibiotics with ≥50% testing rate
4. Specify how many results to display in each graph
5. The tool will:
   - Analyze the data
   - Show interactive graphs
   - Export results to an Excel file

## Output

The tool generates:
1. **Single Antibiotics Analysis**:
   - Susceptibility percentage
   - Confidence intervals
   - Counts of S/R results
   - Testing percentage

2. **Combination Analysis**:
   - Pairwise effectiveness
   - Global effectiveness
   - Confidence intervals for both metrics
   - Testing coverage

3. **Visualizations**:
   - Top single antibiotics
   - Top antibiotic combinations

4. **Excel Export**:
   - All results in separate sheets
   - Graphs embedded in the Excel file

## Statistical Methods

- **Effectiveness calculation**: Percentage of cases where at least one antibiotic in the pair shows susceptibility
- **Confidence intervals**: Calculated using beta distribution for binomial proportions
- **Global effectiveness**: Considers all cases (treating `N` as `R` for combinations)

## Requirements

- Python 3.7+
- Required packages: pandas, numpy, scipy, matplotlib, openpyxl
- For Google Colab: no additional setup needed
