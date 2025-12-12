## Website Link> https://wizardbrew.github.io/Advance-Email-Excel2Outlook-Python/

#Advance-Email-Excel2Outlook 
# VIdeo Guide at the end.
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

     <img width="1920" height="1080" alt="img1" src="https://github.com/user-attachments/assets/8a3ff7f6-75ee-4d97-b300-bba0945465d1" />
---

<img width="1911" height="1080" alt="img2" src="https://github.com/user-attachments/assets/89180bf7-1d1c-4b8e-ad65-9fde034bd9c0" />

---

<img width="1920" height="1080" alt="img3" src="https://github.com/user-attachments/assets/dc6e62da-b29b-4341-bcbc-cbfe95ea48bf" />

---

<img width="1920" height="1080" alt="img4" src="https://github.com/user-attachments/assets/32b0c29c-7926-4688-aaea-0edb147ccf7c" />


---
#Excel File

<img width="1418" height="996" alt="image" src="https://github.com/user-attachments/assets/647e69f3-06bc-4bed-bf0c-44810a26ee80" />


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

##Click Video > 
[![Watch the video](https://img.youtube.com/vi/PXqBLC794Dk/0.jpg)](https://youtu.be/PXqBLC794Dk?si=e7yBiHeKd1LSBRuO)


##############################################################
#                  END OF SOP DOCUMENTATION                  #
##############################################################
