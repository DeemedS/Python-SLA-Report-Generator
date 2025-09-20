import sys
import openpyxl
import os
from datetime import datetime, timedelta
from conn import get_connection
from proc_time_of_day_daily import generate_proc_time_of_day_daily
from proc_time_of_day_weekly import generate_proc_time_of_day_weekly
from proc_sla_daily import generate_sla_daily
from proc_sla_weekly import generate_sla_weekly
from proc_sla_with_cbi_daily import generate_sla_with_cbi_daily
from proc_sla_with_cbi_weekly import generate_sla_with_cbi_weekly
from proc_sla_wo_cbi_daily import generate_sla_wo_cbi_daily
from proc_sla_wo_cbi_weekly import generate_sla_wo_cbi_weekly
from proc_transaction_amount_monthly import generate_transaction_amount_monthly
from proc_pay_cash_amt_monthly import generate_pay_cash_amount_monthly
from proc_total_denom_cbi_monthly import generate_total_denom_cbi_monthly
from proc_total_per_cash_bill_monthly import generate_total_per_cash_bill_monthly


from init_log import init_logger

logger = init_logger(log_dir="logs/txn_analysis", log_name="txn_analysis")      

def main(start_date: str, end_date: str, year_start: str, year: str) -> None:
    logger.info(f"Processing data from {start_date} to {end_date} with year start {year_start}")

    # Convert to datetime for safety
    try:
        start_dt = datetime.strptime(start_date, "%Y-%m-%d")
        end_dt = datetime.strptime(end_date, "%Y-%m-%d")
    except ValueError:
        print("Invalid date format. Use YYYY-MM-DD.")
        return
    
    # Run the stored procedure for generating the data
    try:
        # Create DB connection
        conn = get_connection()
        cursor = conn.cursor()

        # Run stored procedure
        logger.info("Running stored procedure...")

        cursor.execute("EXEC [dbo].[sp_txn_analysis] @START_DT = ?, @END_DT = ?, @YEAR_START = ?", start_dt, end_dt, year_start)
        conn.commit()
        cursor.close()
        conn.close()

        logger.info("Stored procedure executed successfully.")
    except Exception as e:
        logger.error(f"Error running stored procedure: {e}")
        return
    
    #create a excel workbook
    wb = openpyxl.Workbook()

    # Remove the default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    # sheets and corresponding generator functions
    sheets = [
        ("TIMEOFDAY_TXNCOUNT_DAILY", generate_proc_time_of_day_daily, "daily"),
        ("TIMEOFDAY_TXNCOUNT_WEEKLY", generate_proc_time_of_day_weekly, "weekly"),
        ("SLA_DAILY", generate_sla_daily, "daily"),
        ("SLA_WEEKLY", generate_sla_weekly, "weekly"),
        ("SLA_WITH_CBI_DAILY", generate_sla_with_cbi_daily, "daily"),
        ("SLA_WITH_CBI_WEEKLY", generate_sla_with_cbi_weekly, "weekly"),
        ("SLA_WO_CBI_DAILY", generate_sla_wo_cbi_daily, "daily"),
        ("SLA_WO_CBI_WEEKLY", generate_sla_wo_cbi_weekly, "weekly"),
        ("TRANSACTION_AMOUNT", generate_transaction_amount_monthly, "monthly"),
        ("PAY_CASH_AMOUNT - TXN MASTER", generate_pay_cash_amount_monthly, "monthly"),
        ("TOTAL_DENOMINATION - CBI", generate_total_denom_cbi_monthly, "monthly"),
        ("TOTAL_PER_CASH_BILL - CBI", generate_total_per_cash_bill_monthly, "monthly"),
    ]

    # Generate each sheet
    for sheet_name, func, categ in sheets:
        sheet = wb.create_sheet(title=sheet_name)
        if categ in ("daily", "weekly"):
            status = func(sheet, end_date, logger)
        else:
            status = func(sheet, end_date, logger, year_start)
        if status is False:
            logger.error(f"Error generating {sheet_name} sheet. Excel file will not be saved.")
            return
            
    #base folder
    base_folder = r'F:\DW_OneDrive\OneDrive - BTI Payments\DW_STAGING - BTI_DW\Transaction Analysis'
    month_name = end_dt.strftime("%B") 
    year_folder = os.path.join(base_folder, str(year))
    month_folder = os.path.join(year_folder, month_name)
    
    if not os.path.exists(month_folder):
        os.makedirs(month_folder)



    # Save the workbook
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(month_folder, f"{month_name}_txn_analysis_{timestamp}.xlsx")
    wb.save(output_file)
    logger.info(f"Excel file saved as {output_file}")

if __name__ == "__main__":
    if len(sys.argv) > 3:
        start_date = sys.argv[1]
        end_date = sys.argv[2]
        
    else:
        print("Please provide start date and end date in YYYY-MM-DD format.")
        print("Example: python main.py 2025-01-01 2025-03-31 2025")
        sys.exit(1)

    # Convert start_date string to datetime to extract year
    start_dt_obj = datetime.strptime(start_date, "%Y-%m-%d")
    year_start = '2025-06-01' if start_dt_obj.year == 2025 else f"{start_dt_obj.year}-01-01"
    
    year = start_dt_obj.year

    main(start_date, end_date, year_start, year)
