# Automated Options Analyzer (2021)
# Automated Option Analyzer (2021)

## Introduction

The Automated Option Analyzer is a sophisticated tool built using Python to automate the complex analysis of stock options. By leveraging powerful libraries such as NSePY, XLWings, and openpyxl, this analyzer streamlines the process of extracting, processing, and visualizing option data.

## Configuration and Setup

### 1. **Data Source & Libraries**:

- The analyzer employs the `nsepython` and `nse` libraries to fetch real-time stock and option prices.
- Execute the main script `option_analyzer.py` to initiate the analysis.

### 2. **Required Files**:

- `sharelist.txt`: Contains a list of stocks to be analyzed. Input stocks row-wise, with each row containing the share name and its code, separated by a comma (`,`). Ensure the share name is first, followed by its code.
  
- `config.txt`: Houses three critical parameters:
  1. `current_month`: Set to 0 to extract the current month's data or 1 for the next month.
  2. `how_far`: Determines the depth of option price analysis.
  3. `margin`: Utilized in calculating the Internal Rate of Return (IRR).

    **Important Configuration Notes**:
    - Format: `Variable name = variable values`
    - Sequence: Current Month > how far > Margin.
    - Restrictions:
      - `current_month`: Can only be 0 or 1.
      - `how_far` & `margin`: Must be integer values between 0 and 100.
    - Use a hashtag (`#`) for comments.
  
- `Holidays.xlsx`: Enlists the holidays for the ongoing year. Please maintain the existing structure if adding more dates.

### 3. **Before Execution**:

- Ensure all aforementioned files are in the same directory as the main analyzer script.
- On the last Thursday of the month, set `days_left` to 1 (instead of 0) to prevent calculation inaccuracies.
- If updating the stock list in `sharelist.txt`, please delete the previously generated `3-month high low` excel file, `output.xlsx`, and `errors.log` before re-running the analyzer.
- Preserve any code comments by keeping their preceding `#`.

### 4. **Output Files**:

- `3month [timestamp] excel file`: Provides a three-month overview of the highest and lowest stock prices.
  
- `Output.xlsx`: Contains all the raw data pertaining to the shares.

- `Output with [timestamp] excel file`: Represents the final data with essential color-coding to highlight significant insights.

- `Errors with [timestamp].log`: Logs any issues or errors encountered during execution.

### 5. **Color Coding Guide**:

- **Prices (Columns B & C)**:
  - 1-2% change: Bluish Green
  - 2-4% increase: Light Green
  - >4% increase: Green
  - -1 to -2% decrease: Light Red
  - -2 to -4% decrease: Red
  - >-4% decrease: Dark Red

- **RATIO Column**:
  - Ratio > 1.20: Green
  - Ratio < 0.80: Red
  
- **IRRPA (Column K)**:
  - >200: Green
  - Between 150 & 200: Yellow

- **IRRPA (Column K)**:
  - >200: Green
  - Between 150 & 200: Yellow
  
- **Low side and High side (Columns N & O)**:
  - Value > 9: Green
  - Between 6 & 9: Yellow

- **F SELLPE of and F SELL PE @ (Columns R & S)**:
  - Data not found: Light Brown
  - `turns=1`: No Color
  - `turns=2` & `turns=3`: Orangish Red
  - `turns=4`: Red
  
- **F SELLCE of and F SELL CE @ (Columns V & W)**:
  - Data not found: Light Brown
  - `turns=1`: No Color
  - `turns=2` & `turns=3`: Orangish Red
  - `turns=4`: Red
  
- **F IRR PA AND F IRR PA1 (Columns U & Y)**:
  - Value >= 48%: Green
  - Value >= 36% but < 48%: Light Green
  - Value >= 24% but < 36%: Yellow
  - Value < 24%: No Color

---

These color codings serve as a quick visual reference, allowing users to instantly recognize patterns, shifts, and significant data points within the analysis. By understanding these codings, one can rapidly interpret the insights presented in the output files.

## Conclusion

The Automated Option Analyzer provides a robust and intuitive solution for those seeking to demystify the intricacies of stock options. Through its systematic extraction, processing, and visually guided representation, the analyzer streamlines the decision-making process for investors and analysts alike. The distinctive color-coding system enhances data interpretability, ensuring that significant trends and data points are effortlessly identifiable. As stock markets continue to evolve, tools like this will be indispensable for professionals aiming to stay ahead of the curve, making well-informed, data-driven decisions.
