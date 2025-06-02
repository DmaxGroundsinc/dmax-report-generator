
import pandas as pd
from docx import Document
import os

def fill_dmax_report(excel_path, template_path, output_path):
    xls = pd.ExcelFile(excel_path)

    property_df = xls.parse('Property Overview', index_col=0).dropna()
    cash_df = xls.parse('Cash Offer', index_col=0).dropna()
    finance_df = xls.parse('Seller Financing', index_col=0).dropna()
    final_df = xls.parse('Final Summary', index_col=0).dropna()

    placeholder_map = {
        "{{PropertyAddress}}": property_df.get("Input", {}).get("Property Address", ""),
        "{{ListingPrice}}": property_df.get("Input", {}).get("Listing Price", ""),
        "{{SquareFootage}}": property_df.get("Input", {}).get("Square Footage", ""),
        "{{LotSize}}": property_df.get("Input", {}).get("Lot Size", ""),
        "{{PropertyType}}": property_df.get("Input", {}).get("Property Type", ""),
        "{{AgentName}}": property_df.get("Input", {}).get("Agent Name", ""),
        "{{AgentEmail}}": property_df.get("Input", {}).get("Agent Email", ""),
        "{{AgentPhone}}": property_df.get("Input", {}).get("Agent Phone", ""),
        "{{Brokerage}}": property_df.get("Input", {}).get("Brokerage", ""),
        "{{CashOfferPrice}}": cash_df.get("Value", {}).get("Offer Price", ""),
        "{{CashClosingCosts}}": cash_df.get("Value", {}).get("Closing Costs", ""),
        "{{CashRepairs}}": cash_df.get("Value", {}).get("Repairs", ""),
        "{{CashCommission}}": cash_df.get("Value", {}).get("Commission (3%)", ""),
        "{{CashTotalInvestment}}": cash_df.get("Value", {}).get("Total Investment", ""),
        "{{CashMonthlyRent}}": cash_df.get("Value", {}).get("Monthly Rent", ""),
        "{{CashNetMonthlyIncome}}": cash_df.get("Value", {}).get("Net Monthly Income", ""),
        "{{CashAnnualNetIncome}}": cash_df.get("Value", {}).get("Annual Net Income", ""),
        "{{CashCoCReturn}}": cash_df.get("Value", {}).get("Cash-on-Cash Return", ""),
        "{{FinancePurchasePrice}}": finance_df.get("Value", {}).get("Purchase Price", ""),
        "{{FinanceDownPayment}}": finance_df.get("Value", {}).get("Down Payment (10%)", ""),
        "{{FinanceClosingCommission}}": finance_df.get("Value", {}).get("Closing + Commission", ""),
        "{{FinanceTotalCashIn}}": finance_df.get("Value", {}).get("Total Cash In", ""),
        "{{FinanceMonthlyPayment}}": finance_df.get("Value", {}).get("Monthly Payment to Seller", ""),
        "{{FinancePITI}}": finance_df.get("Value", {}).get("PITI", ""),
        "{{FinanceTotalMonthlyOut}}": finance_df.get("Value", {}).get("Total Monthly Out", ""),
        "{{FinanceNetMonthlyIncome}}": finance_df.get("Value", {}).get("Net Monthly Income", ""),
        "{{FinanceAnnualNetIncome}}": finance_df.get("Value", {}).get("Annual Net Income", ""),
        "{{FinanceCoCReturn}}": finance_df.get("Value", {}).get("Cash-on-Cash Return", ""),
        "{{SuggestedLockupPrice}}": final_df.get("Value", {}).get("Suggested Lock-up Price", ""),
        "{{AssignmentFee}}": final_df.get("Value", {}).get("Assignment Fee", ""),
        "{{DoubleCloseMargin}}": final_df.get("Value", {}).get("Double Close Margin", ""),
        "{{BestInvestorProfile}}": final_df.get("Value", {}).get("Best Investor Profile", ""),
    }

    doc = Document(template_path)

    for para in doc.paragraphs:
        for key, value in placeholder_map.items():
            if key in para.text:
                para.text = para.text.replace(key, str(value))

    doc.save(output_path)
    print(f"Report saved to: {output_path}")

# Example usage:
# fill_dmax_report("DMAX_Master_Spreadsheet_Model_With_EntryFormatting.xlsx", 
#                  "DMAX_Branded_Template_With_Placeholders.docx", 
#                  "DMAX_Filled_Deal_Report.docx")
