# 💰 Excel Pricing Automation Model

## 🧩 Overview
This project presents a **pricing automation tool developed in Excel with VBA**, designed to help a service company define its prices more accurately and consistently.  
The tool centralizes all financial and operational data — such as employee costs, vehicle expenses, and fixed costs — to calculate the service price automatically based on the selected resources and desired profit margin.

---

## ⭐ Methodology: STAR

### **Situation**
The company had difficulty determining the correct pricing for its services.  
Costs were dispersed across multiple spreadsheets, making it hard to calculate labor, vehicle, and fixed expenses precisely.  
This led to **inconsistent pricing** and reduced profitability on certain contracts.

---

### **Task**
My goal was to **develop an integrated pricing model** that would:
- Consolidate all cost sources in one structured file.  
- Automate price calculation based on parameters chosen by the user.  
- Allow the company to **quickly simulate** different project scenarios with varying profit margins and tax rates.  

---

### **Action**
I built a complete Excel workbook organized in **five interconnected sheets**, using formulas and **VBA automation** to ensure real-time updates between them:

1. **🧭 Home Page** – Instructions and user guidance.  
2. **👥 Collaborators** – Table with all employees, their salaries, and areas of expertise.  
3. **🚗 Fleet** – List of company vehicles, including maintenance costs, IPVA, tolls, mileage (km/L), and related expenses.  
4. **🏢 Fixed Costs** – Monthly fixed costs such as rent, utilities, and administrative expenses.  
5. **📊 Pricing** – Interactive pricing calculator:
   - Selects employees involved in the service and their working hours.  
   - Selects vehicles used, fuel cost, and distance.  
   - Defines profit margin and applicable taxes.  
   - Pulls all related data automatically from the previous sheets.

💡 **Automation Example:**  
When a new employee is added in the *Collaborators* sheet, the name automatically appears in the selection list of the *Pricing* sheet — no manual update needed.  

---

### **Result**
The final tool allowed the company to:
- **Standardize** its pricing process;  
- **Save time** by eliminating manual cost consolidation;  
- **Simulate different profit margins and tax scenarios** in seconds;  
- Improve **decision-making** and confidence in the values presented to clients.  

---

## 🔒 Confidentiality
Due to **confidentiality and non-disclosure agreements**, the original spreadsheet cannot be shared.  
All images in this repository use **fictional data** and **a neutral design** — company names, logos, and financial values have been fully anonymized.  
Only the **visual structure and automation logic** are represented here for demonstration purposes.

---

## 🖼️ Previews
| Example Screens | Description |
|------------------|-------------|
| ![Home Page](images/home_preview.png) | Instructions and project overview |
| ![Pricing Sheet](images/pricing_preview.png) | Automated pricing interface with linked data |

---

## 🧠 Tools & Techniques
- **Microsoft Excel (Advanced)**  
- **VBA Automation (Macros)**  
- **Data Linking & Consolidation**  
- **Financial Modeling & Cost Structure Design**
