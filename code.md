import pandas as pd
import os
from datetime import datetime
import win32com.client
import time
 
# ============================================================
# CONFIG
# ============================================================
CSV_FILE_PATH = r"D:\OneDrive - TATA MOTORS LTD\Desktop\TMS_Report\Report.csv"
MANAGER_EMAIL = "nt922402.ttl@tatamotors.com; samirt.ttl@tatamotors.com"
CC_EMAILS = "dj922785.ttl@tatamotors.com; ss922682.ttl@tatamotors.com; p0001882.ttl@tatamotors.com"
YOUR_EMAIL = "rm925712.ttl@tatamotors.com"
DATE_FILTER = "today" 
 
# ============================================================
# CSV LOAD & FILTER
# ============================================================
def load_and_filter_csv(file_path):
    if not os.path.exists(file_path):
        print("❌ CSV file not send")
        return None, None, None
 
    df = pd.read_csv(file_path)
    print(f"✅ Total tickets loaded: {len(df)}")
 
    # Solved
    solved_df = df[df['Ticket Status'] == 'Solved'].copy()
 
    # Date filter
    if DATE_FILTER == "today":
        today = datetime.now().strftime("%d-%b-%Y")
        solved_df = solved_df[
            solved_df['Resolved Time'].str.contains(today, na=False)
        ]
 
    # Unclaimed (Raised)
    unclaimed_df = df[df['Ticket Status'] == 'Raised'].copy()
 
    # Claimed but Not Solved
    claimed_not_solved_df = df[df['Ticket Status'] == 'Claimed'].copy()
 
    print(f"✅ Solved: {len(solved_df)}, Unclaimed: {len(unclaimed_df)}, Claimed: {len(claimed_not_solved_df)}")
    return solved_df, unclaimed_df, claimed_not_solved_df
 
# ============================================================
# COUNT SOLVED PER PERSON
# ============================================================
def count_tickets_per_person(solved_df):
    person_count = {}
    for _, row in solved_df.iterrows():
        support_persons = str(row['Support Persons'])
        if support_persons == 'nan' or support_persons.strip() == '':
            continue
        persons = [p.strip() for p in support_persons.split(',')]
        for person in persons:
            if person:
                person_count[person] = person_count.get(person, 0) + 1
    return dict(sorted(person_count.items(), key=lambda x: x[1], reverse=True))
 
# ============================================================
# COUNT CLAIMED PER PERSON
# ============================================================
def count_claimed_per_person(df):
    """Claimed tickets per support person"""
    person_count = {}
    claimed_df = df[df['Ticket Status'] == 'Claimed'].copy()
    for _, row in claimed_df.iterrows():
        support_persons = str(row['Support Persons'])
        if support_persons == 'nan' or support_persons.strip() == '':
            continue
        persons = [p.strip() for p in support_persons.split(',')]
        for person in persons:
            if person:
                person_count[person] = person_count.get(person, 0) + 1
    return person_count
 
 
 
# ============================================================
# CREATE EXCEL REPORT
# ============================================================
def create_report_excel(report_path, solved_df, unclaimed_df, claimed_not_solved_df,
                         person_solved, person_claimed):
 
    
    all_persons = set(list(person_solved.keys()) +
                      list(person_claimed.keys()))
 
    summary_data = []
    for person in all_persons:
        summary_data.append({
            'Team Member': person,
            'Tickets Solved': person_solved.get(person, 0),
            'Tickets Claimed': person_claimed.get(person, 0),
        })
 
    summary_df = pd.DataFrame(summary_data)
    summary_df = summary_df.sort_values('Tickets Solved', ascending=False).reset_index(drop=True)
    summary_df.insert(0, 'Rank', range(1, len(summary_df) + 1))
 
    with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        solved_df.to_excel(writer, sheet_name='Solved Tickets', index=False)
        unclaimed_df.to_excel(writer, sheet_name='Unclaimed Tickets', index=False)
        claimed_not_solved_df.to_excel(writer, sheet_name='Claimed Not Solved', index=False)
 
    print("✅ Excel report created successfully")
    return summary_df
 
# ============================================================
# SEND EMAIL
# ============================================================
def send_outlook_email(report_path, solved_df, unclaimed_df, claimed_not_solved_df,
                        person_solved, person_claimed, summary_df):
 
    today_str = datetime.now().strftime("%d %B %Y")
    total_solved = len(solved_df)
    total_unclaimed = len(unclaimed_df)
    total_claimed_not_solved = len(claimed_not_solved_df)
 
    # HTML Table - 4 columns
    table_rows = ""
    for _, row in summary_df.iterrows():
        table_rows += f"""
        <tr>
            <td style="border:1px solid #ddd; padding:8px; text-align:center;">{int(row['Rank'])}</td>
            <td style="border:1px solid #ddd; padding:8px;">{row['Team Member']}</td>
            <td style="border:1px solid #ddd; padding:8px; text-align:center; color:green; font-weight:bold;">{int(row['Tickets Solved'])}</td>
            <td style="border:1px solid #ddd; padding:8px; text-align:center; color:#e67e22; font-weight:bold;">{int(row['Tickets Claimed'])}</td>
        </tr>"""
 
    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <p>Dear Sir,</p>
 
        <h3 style="color:#1a5276;">📊 Ticket Summary - {today_str}</h3>
        <table style="border-collapse:collapse;">
            <tr>
                <td style="padding:6px 20px 6px 0;"><b>✅ Total Solved:</b></td>
                <td style="color:green; font-weight:bold;">{total_solved}</td>
            </tr>
            <tr>
                <td style="padding:6px 20px 6px 0;"><b>⏳ Unclaimed (Raised):</b></td>
                <td style="color:red; font-weight:bold;">{total_unclaimed}</td>
            </tr>
            <tr>
                <td style="padding:6px 20px 6px 0;"><b>🔄 Claimed but Not Solved:</b></td>
                <td style="color:#e67e22; font-weight:bold;">{total_claimed_not_solved}</td>
            </tr>
        </table>
 
        <br>
        <h3 style="color:#1a5276;">👥 Team-wise Ticket Count</h3>
        <table style="border-collapse:collapse; min-width:500px;">
            <thead>
                <tr style="background-color:#1a5276; color:white;">
                    <th style="border:1px solid #ddd; padding:10px;">Rank</th>
                    <th style="border:1px solid #ddd; padding:10px;">Team Member</th>
                    <th style="border:1px solid #ddd; padding:10px;">Solved</th>
                    <th style="border:1px solid #ddd; padding:10px;">Claimed</th>
                </tr>
            </thead>
            <tbody>
                {table_rows}
            </tbody>
        </table>
 
        <br>
        <p>Detailed report is attached for your reference.</p>
        <p>Regards,<br><b>Rohini Malekar</b></p>
    </body>
    </html>
    """
 
    try:
        outlook = win32com.client.Dispatch('outlook.application')
        time.sleep(5)
        mail = outlook.CreateItem(0)
        mail.To = MANAGER_EMAIL
        mail.CC = CC_EMAILS
        mail.Subject = f"TMS Daily Report | {today_str} | Solved: {total_solved} | Unclaimed: {total_unclaimed}"
        mail.HTMLBody = html_body
        if os.path.exists(CSV_FILE_PATH):
            mail.Attachments.Add(os.path.abspath(CSV_FILE_PATH))
        mail.Send()
        print("✅ Email sent successfully")
    except Exception as e:
        print("❌ Email failed:", e)
 
# ============================================================
# MAIN
# ============================================================
def main():
    print("=" * 60)
    print("  PFIRST TMS - Daily Ticket Report Automation")
    print("=" * 60)
 
    df_full = pd.read_csv(CSV_FILE_PATH) if os.path.exists(CSV_FILE_PATH) else None
    if df_full is None:
        print("❌ CSV file not found")
        return
 
    solved_df, unclaimed_df, claimed_not_solved_df = load_and_filter_csv(CSV_FILE_PATH)
    if solved_df is None:
        return
 
    # Counts
    person_solved = count_tickets_per_person(solved_df)
    person_claimed = count_claimed_per_person(df_full)
 
    # Report
    today_file = datetime.now().strftime("%d-%m-%Y")
    report_path = f"C:\\TMS_Reports\\TMS_Report_{today_file}.xlsx"
    os.makedirs(os.path.dirname(report_path), exist_ok=True)
 
    summary_df = create_report_excel(
        report_path, solved_df, unclaimed_df, claimed_not_solved_df,
        person_solved, person_claimed
    )
 
    print("\n👥 Team Ticket Count:")
    print(summary_df.to_string(index=False))
 
    send_outlook_email(
        report_path, solved_df, unclaimed_df, claimed_not_solved_df,
        person_solved, person_claimed, summary_df
    )
 
    print("\n✅ Done!")
    print("=" * 60)
 
if __name__ == "__main__":
    main()
 