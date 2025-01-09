# Automating-Excel-tasks-with-openpyxl

# Healthcare Stroke Data Analysis with Excel Automation  

## Project Overview  
This project focuses on automating the analysis of a **Healthcare Stroke Dataset** using Python libraries such as `openpyxl` and `pandas`. The automation enables efficient data manipulation, visualization, and reporting, reducing manual effort and improving scalability for larger datasets.  

## Key Features  
1. **High-Risk Patient Filtering**:  
   - Filters and saves stroke patients into a separate workbook for focused analysis.  

2. **Summary Report Creation**:  
   - Generates a report with key metrics like average age, BMI, glucose levels, and stroke counts.  

3. **Categorical Data Count**:  
   - Counts and organizes categorical data such as gender and smoking status into a structured sheet.  

4. **Critical Case Highlighting**:  
   - Applies conditional formatting to highlight critical cases based on glucose levels and BMI thresholds.  

5. **Visualization with Charts**:  
   - Creates a bar chart to visualize stroke counts by gender directly in Excel.  

## Tools and Libraries Used  
- **Python**  
  - `openpyxl`: For Excel automation including data manipulation, formatting, and charting.  
  - `pandas`: For efficient data handling and analysis.  
- **Excel**: For storing and visualizing automated outputs.  

## How to Run the Project  
1. **Prerequisites**:  
   - Install Python (version 3.8 or later).  
   - Install the required libraries using pip:  
     ```bash
     pip install openpyxl pandas
     ```  

2. **Dataset**:  
   - Place the `healthcare-stroke-data.xlsx` file in the project directory.  

3. **Execution**:  
   - Run the Python script to perform the analysis:  
     ```bash
     python automate_stroke_analysis.py
     ```  

4. **Outputs**:  
   - `high_risk_patients.xlsx`: A filtered workbook containing high-risk stroke patients.  
   - `automated_healthcare_analysis.xlsx`: A comprehensive workbook with summary reports, category counts, highlighted critical cases, and visualizations.  

## Project Files  
- `automate_stroke_analysis.py`: Python script for automating the tasks.  
- `healthcare-stroke-data.xlsx`: Input dataset.  
- `high_risk_patients.xlsx`: Filtered output.  
- `automated_healthcare_analysis.xlsx`: Final output workbook with all tasks completed.  

## Tasks Performed  
1. **Filter High-Risk Patients**:  
   - Extracted patients with stroke conditions (`stroke == 1`).  
   - Saved results to a new workbook.  

2. **Generate Summary Report**:  
   - Calculated average age, BMI, glucose levels, and total patient counts.  

3. **Count Categorical Data**:  
   - Counted values for gender and smoking status.  

4. **Highlight Critical Cases**:  
   - Used conditional formatting to highlight patients with glucose > 200 or BMI > 30.  

5. **Bar Chart Visualization**:  
   - Created a bar chart for stroke counts by gender.  



## Conclusion  
This project demonstrates how Python can automate Excel tasks, making data analysis faster, more accurate, and scalable. It provides a foundation for handling healthcare data and generating actionable insights.  

## Acknowledgments  
- **Entri Elevate**: For the 75-day Data Analysis Challenge.  
- Python community and open-source contributors for their excellent libraries.  
