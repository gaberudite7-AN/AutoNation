import pdfplumber
import pandas as pd
import re

# Containers
structured_data = []
dealer_info = {}

# Keywords to identify fundamentals
fundamental_keywords = [
    "Facility", "Dealer Naming", "EV Readiness", "Facility Conditions",
    "GM Exclusive", "Audi Experience", "Sales Reporting", "Financial", "Training"
]

# Open the PDF
with pdfplumber.open(r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Triggers\PDF_Convert\PDF\data.pdf") as pdf:
    for page_num, page in enumerate(pdf.pages, start=1):
        text = page.extract_text()
        if not text:
            continue
        
        lines = text.split('\n')
        for line in lines:
            # Debug: print each line
            print(f"[Page {page_num}] {line}")

            # Dealer Code
            if "Dealer Code:" in line:
                match = re.search(r"Dealer Code:\s*(\S+)", line)
                if match:
                    dealer_info["Dealer Code"] = match.group(1)

            # Dealer Name (example here hardcoded for Audi Bellevue, adapt if needed)
            if "Audi Bellevue" in line:
                dealer_info["Dealer Name"] = "Audi Bellevue"

            # Extract fundamental rows
            if any(keyword in line for keyword in fundamental_keywords):
                structured_data.append({"Fundamental": line.strip()})

# Convert to DataFrames
df_fundamentals = pd.DataFrame(structured_data)
df_dealer = pd.DataFrame([dealer_info])

# Save to Excel
with pd.ExcelWriter(r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Triggers\PDF_Convert\Excel\structured_output.xlsx") as writer:
    df_dealer.to_excel(writer, sheet_name="Dealer Info", index=False)
    df_fundamentals.to_excel(writer, sheet_name="Fundamentals", index=False)

# Save to HTML
html_content = "<html><body>"
html_content += "<h2>Dealer Info</h2>"
html_content += df_dealer.to_html(index=False)
html_content += "<h2>Business Fundamentals</h2>"
html_content += df_fundamentals.to_html(index=False)
html_content += "</body></html>"

with open(r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Triggers\PDF_Convert\Excel\structured_output.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("✅ Structured data extracted and saved to 'structured_output.xlsx' and 'structured_output.html'.")
