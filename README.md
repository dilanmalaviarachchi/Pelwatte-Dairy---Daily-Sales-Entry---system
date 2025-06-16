

# **Pelwatte Dairy - Daily Sales Report System**

## **Overview**
This **Streamlit** application is designed to simplify the daily sales reporting process for **Pelwatte Dairy Industries Limited**. It allows sales representatives to enter daily sales data, which is then automatically formatted and saved into the company's existing Excel template while preserving all original formulas, headers, and structure.

## **Key Features**
✅ **Excel Template Compatibility**  
- Works with the existing `wattala RD Vehicle Allowance 2025.xlsx` template  
- Preserves all formatting, formulas, and structure  

✅ **Daily Data Entry**  
- Simple form for entering:  
  - Sales quantities (FCMP, Yoghurt, Ice-Cream, Butter, etc.)  
  - RD Allowance  
  - Call metrics (Scheduled Calls, Calls Made, Product Calls)  
- Auto-fills today's date  

✅ **Data Management**  
- Maintains all historical entries for the month  
- Displays current month's data in an interactive table  

✅ **Automatic Excel Export**  
- Generates an updated Excel file with:  
  - Original headers & footers  
  - All formulas intact  
  - New data in the correct format  
- Downloadable with a timestamped filename  

✅ **User-Friendly Interface**  
- Step-by-step workflow  
- Input validation to prevent errors  
- Clear success/error messages  

## **How to Use**
1. **Upload Template**  
   - Start by uploading your existing `wattala RD Vehicle Allowance 2025.xlsx` file.  

2. **Enter Daily Data**  
   - Fill in the form with today's sales and call metrics.  
   - Click **"Submit Today's Report"** to save.  

3. **View & Export**  
   - Check the current month's data in the table.  
   - Download the updated Excel file anytime.  

## **Installation & Setup**
1. **Install Python (if not already installed)**  
   - Download from [python.org](https://www.python.org/downloads/)  

2. **Install Required Libraries**  
   ```sh
   pip install streamlit pandas openpyxl
   ```

3. **Run the App**  
   ```sh
   streamlit run pelwatte_daily_entry.py
   ```

4. **Open in Browser**  
   - The app will automatically open at `http://localhost:8501`  


## **Why Use This?**
✔ **Saves Time** – No manual Excel editing  
✔ **Reduces Errors** – Structured data entry  
✔ **Keeps Consistency** – Maintains original Excel format  
✔ **Easy to Use** – No technical skills required  

## **Future Improvements**
- **Cloud Storage** – Auto-save reports to Google Drive/Dropbox  
- **Multi-User Support** – Track different sales reps  
- **Dashboard Analytics** – Visualize monthly sales trends  

