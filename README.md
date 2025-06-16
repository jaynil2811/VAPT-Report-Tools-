# VAPT-Report-Tools-
Here , VAPT Reports Like , Feed the Nmap Result (TXT) file to Generate EXCEL Report , NESSUS Report to Final Report 
🛠️ Security Tool Suite 🛠️
Welcome to the Security Tool Suite repository! This suite includes four powerful tools designed to automate the generation of security reports from various sources such as Nmap, Nessus, and Annexure data. The tools output clean, actionable results, making it easier for security professionals to assess vulnerabilities and generate comprehensive reports.

🛠️ Tools in this Suite:
1️⃣ Nmap Result to Excel Data 🖥️
This tool processes the output from Nmap and organizes it into an Excel file. The file is structured to allow for easy analysis and reporting.

Features:

Extracts data from Nmap results.

Feeds the data into an Excel sheet.

Formats the results into Annexure 2 for easy viewing and analysis.

2️⃣ Nessus Report to Filtered Data & Final Report 📊
This tool takes Nessus CSV reports, filters them by severity, and compares the data with a template database. It generates a final, filtered report, showing only the most critical vulnerabilities.

Features:

Imports Nessus CSV data.

Filters data by severity level.

Compares the filtered data with a template database.

Generates the final company report with actionable insights.

3️⃣ Add Data From the Annexure 2 🗂️
This tool extracts and fills in missing data from Annexure 2, such as HTTP, FTP, and Unknown protocol data.

Features:

Completes missing Annexure 2 data.

Fills in protocols like HTTP, FTP, and others.

Ensures data completeness for final report generation.

4️⃣ Final Report Generator 📑
The final tool consolidates all the processed data from previous steps and generates a clean, professional security report for the company.

Features:

Generates a final, comprehensive security report.

Consolidates results from Nmap, Nessus, and Annexure data.

Outputs a polished, ready-to-share report for stakeholders.

🚀 Installation & Setup
Process :
git clone https://github.com/jaynil2811/VAPT-Report-Tools-/
pip install -r requirements.txt
Run the Exe or python as you Need Both are Same ! 

Before you start using the tools, make sure to follow these instructions to get everything set up:

1. Clone the Repository
First, clone the repo to your local machine:

bash
Copy
git clone https://github.com/jaynil2811/VAPT-Report-Tools-/
2. Install Requirements (if you want Run Python File Otherwise Exe are also Availible !)
Install the necessary Python packages and dependencies by using the requirements file:

bash
Copy
pip install -r requirements.txt
3. Download & Install Executables
Ensure that you have all the .exe files and templates needed to run the tools. You can find these in the /exe folder after downloading.

4. Run the Tools
You can run each tool as needed:

For Tool 1 (Nmap to Excel), simply execute the .exe file or script after placing your Nmap result.

For Tool 2 (Nessus Report Filter), input your CSV file and template database.

For Tool 3 (Annexure 2 Data Add), ensure that the Annexure 2 file is provided to fill in missing data.

For Tool 4 (Final Report), generate the final report by combining all the processed data.

💡 How It Works:
Input Data: Start by providing the input files (Nmap results, Nessus CSVs, and Annexure data).

Process Data: The tools filter, extract, and format the data into actionable insights.

Generate Reports: The final step is the creation of a professional security report ready for analysis and action.



📚 Contributing
We welcome contributions to improve these tools! If you have suggestions or want to fix issues, feel free to fork the repo and submit a pull request.

📬 Contact & Support
If you have any questions or need further assistance, feel free to open an issue or reach out!

