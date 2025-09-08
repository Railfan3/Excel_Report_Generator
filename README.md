
# 📊 Excel Report Generator  

## 🎯 Objective  
Generate professional Excel summaries from CSV inputs with pivot tables, charts, and styled outputs.  

---

## 🛠 Tools & Libraries  
- **pandas** → CSV reading, cleaning, summarization  
- **openpyxl** → Excel writing & styling  
- **Tkinter** → File dialogs & simple GUI  
- **matplotlib** → Chart generation  

---

## 🚀 Features  
✔ Load CSV into pandas DataFrame  
✔ Create **pivot tables** for summarized insights  
✔ Generate **charts** (bar, line, pie, etc.) with matplotlib  
✔ Export **styled Excel reports** using openpyxl  
✔ GUI with **file upload & save dialogs** (Tkinter)  
✔ Add **summary statistics** (mean, median, min, max, etc.)  

---

## 📂 Project Structure
excel-report-generator/
│── main.py # Entry point (Tkinter GUI + logic)
│── report_generator.py # Core functions (pandas, openpyxl, charts)
│── requirements.txt # Required Python libraries
│── sample_data.csv # Example input file
│── outputs/ # Generated Excel reports
│── README.md # Project documentation


---

## 📝 Quick Guide  

1. **Load CSV into pandas**
   ```python
   import pandas as pd  
   df = pd.read_csv("data.csv")  
Create pivot tables

pivot = pd.pivot_table(df, index="Category", values="Sales", aggfunc="sum")


Generate charts with matplotlib

import matplotlib.pyplot as plt  
pivot.plot(kind="bar")  
plt.savefig("chart.png")  


Export styled Excel using openpyxl

with pd.ExcelWriter("report.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Raw Data")
    pivot.to_excel(writer, sheet_name="Summary")


Add file dialogs with Tkinter

from tkinter import filedialog  
file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])


Add summary stats

summary = df.describe(include="all")

▶️ How to Run

Clone the repository

git clone https://github.com/yourusername/excel-report-generator.git
cd excel-report-generator


Install dependencies

pip install -r requirements.txt


Run the app

python main.py

📌 Example Output

Raw Data Sheet → Original CSV data

Summary Sheet → Pivot tables + statistics

Charts Sheet → Visualizations (bar, line, pie)

Styled Excel Report → Ready for analysis

✅ Future Enhancements

📂 Support multiple CSV uploads

📊 Advanced chart customization

📅 Scheduling automated reports

📈 Export to interactive dashboards

