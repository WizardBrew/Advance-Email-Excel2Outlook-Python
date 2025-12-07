import pandas as pd
import win32com.client as win32
import os
import csv
from datetime import datetime

# Load Excel file
df = pd.read_excel(r"D:\OneDrive\Backedup\Desktop\EmailExcel2OutlookAttachment\emails.xlsx")

# Connect to Outlook
outlook = win32.Dispatch("Outlook.Application")

sent_list = []
failed_list = []
skipped_list = []

total_rows = len(df)
print(f"📊 Starting email process for {total_rows} rows...\n")

stop_process = False

for idx, row in df.iterrows():
    if stop_process:
        break  # stop sending if cancelled earlier

    row_number = idx + 2
    recipient = str(row.get("Email", "")).strip()
    subject = str(row.get("Subject", "")).strip()
    body = str(row.get("Body", "")).strip()
    attachment = str(row.get("AttachmentPath", "")).strip()

    print(f"➡️ Processing row {row_number}/{total_rows}...")

    if not recipient or recipient.lower() == "nan":
        msg = f"Row {row_number} (no email)"
        print(f"⏭️ Skipping {msg}")
        skipped_list.append(msg)
        continue

    try:
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.Body = body

        if attachment and attachment.lower() != "nan":
            clean_path = os.path.normpath(attachment)
            if os.path.isfile(clean_path):
                mail.Attachments.Add(clean_path)
                print(f"📎 Attached file for {recipient}: {clean_path}")
            else:
                print(f"⚠️ Attachment not found for {recipient}: {clean_path}")

        # First 3 emails → Preview mode
        if idx < 3:
            mail.Display()
            print(f"👀 Preview opened for {recipient}. Please review in Outlook.")
            user_input = input("Press Enter to continue, or type 'cancel' to stop: ")
            if user_input.lower() == "cancel":
                print("🛑 Process cancelled by user. No further emails will be sent.")
                stop_process = True
                break
            sent_list.append(recipient)
        else:
            # Rest → Auto-send
            mail.Send()
            print(f"✅ Sent to {recipient}")
            sent_list.append(recipient)

    except Exception as e:
        print(f"❌ Failed to send to {recipient}: {e}")
        failed_list.append(recipient)

# --- Log Report ---
print("\n--- Log Report ---")
print(f"✅ Sent ({len(sent_list)}):")
for r in sent_list:
    print(f"   - {r}")

print(f"\n❌ Failed ({len(failed_list)}):")
for r in failed_list:
    print(f"   - {r}")

print(f"\n⏭️ Skipped ({len(skipped_list)}):")
for r in skipped_list:
    print(f"   - {r}")

print(f"\n📊 Total rows processed: {len(df)}")

# --- Write CSV Log ---
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
log_file = f"email_log_{timestamp}.csv"

with open(log_file, mode="w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Status", "Recipient/Row"])
    for r in sent_list:
        writer.writerow(["Sent", r])
    for r in failed_list:
        writer.writerow(["Failed", r])
    for r in skipped_list:
        writer.writerow(["Skipped", r])

print(f"\n📝 Log saved to: {log_file}")
