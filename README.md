# Advance-Email-Excel2Outlook

##############################################################
#                  EMAIL AUTOMATION SOP                      #
#        Python + Pandas + Outlook COM Integration           #
##############################################################

📊 WORKFLOW OVERVIEW
--------------------------------------------------------------
1. LOAD EXCEL FILE
   - Reads: emails.xlsx
   - Columns: Email | Subject | Body | AttachmentPath
   - Each row = one email

2. ITERATE THROUGH ROWS
   - Sequential processing
   - Skips invalid/missing email addresses

3. CREATE OUTLOOK EMAIL
   - Recipient, Subject, Body populated
   - Attachment added if file exists

4. PREVIEW vs AUTO-SEND
   - First 3 emails → mail.Display()
     * User manually clicks Send
     * Console prompt:
       > Press Enter to continue, or type 'cancel' to stop
     * Cancel halts script immediately
   - Remaining emails → mail.Send() auto-send

5. LOGGING
   - Console progress:
     ➡️ Processing row X/Y
   - End report:
     ✅ Sent list
     ❌ Failed list
     ⏭️ Skipped list
   - CSV log generated:
     email_log_<timestamp>.csv
     Columns: Status | Recipient/Row

     <img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/62e135c0-c824-430c-a75d-ed0122180df7" />

     <img width="1911" height="1080" alt="image" src="https://github.com/user-attachments/assets/079d3570-5625-492c-b99f-3d57855673e8" />

     <img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/c2342c7e-7ff2-4202-a069-9cbfe1de99d8" />

     #Log Genration.
     <img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/c46ac011-449f-4bf9-89ba-14c9d610bfbb" />





--------------------------------------------------------------
⚙️ BEHAVIOR DETAILS
--------------------------------------------------------------
- Skipped Rows: Empty/NaN emails → safely skipped
- Attachments: Added only if file exists
- Preview Mode: First 3 emails open for manual review
- Cancel Option: Stops script instantly
- Auto-Send Mode: Remaining emails sent automatically
- Error Handling: COM/Outlook errors logged as Failed

--------------------------------------------------------------
🛡️ SAFETY & RELIABILITY
--------------------------------------------------------------
- Prevents accidental bulk send (preview first 3)
- Cancel option for user control
- Comprehensive logs for audit & troubleshooting

--------------------------------------------------------------
🚀 QUICK START
--------------------------------------------------------------
$ python email_automation.py

##############################################################
#                  END OF SOP DOCUMENTATION                  #
##############################################################
