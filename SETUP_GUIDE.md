# The BYOD Automation Setup Guide

## The Goal

Automate the entire BYOD process:

1. Google Forms → Auto-sync to Excel
2. Email to Supervisor with clickable Approve/Reject buttons
3. Automatic IT inspection scheduling(2 days from date of approval)
4. QR code generation

---

## Part 1: Google Forms Auto-Sync

**1. Install Packages:**

```bash
pip install gspread google-auth --break-system-packages
```

  <!-->

    its running in a venv so no need for break-systems-packages

</!-->

**2. Set Up Google Cloud (One-time):**

Go to: https://console.cloud.google.com

**Create Project:**

1. Click "New Project"
2. Name: "NITDA BYOD System"
3. Click "Create"

**Enable APIs:**

1. Search "Google Sheets API" → Enable
2. Search "Google Drive API" → Enable

**Create Service Account:**

1. Menu → IAM & Admin → Service Accounts
2. Click "Create Service Account"
3. Name: "BYOD Automation"
4. Click "Create and Continue"
5. Skip optional steps
6. Click "Done"

**Generate Key:**

1. Click on the service account
2. Keys tab → Add Key → Create new key
3. Choose JSON
4. Click "Create"
5. Save as `credentials.json` in your project folder

**3. Share Google Sheet:**

1. Open the JSON file
2. Find `client_email` (should look like: xxx@xxx.iam.gserviceaccount.com)
3. Copy this email
4. Open your Google Sheet
5. Click Share
6. Paste the email
7. Give "Editor" access
8. Click "Send"

**4. Run byod_automate:**

```bash
python byod_automate.py
```

Choose option 1, enter your Google Sheet name(The sheet where responses are stored from the form)

**Result:** Script runs continuously, checking for new responses every 60 seconds!

---

## Part 2: Email Approval Buttons

### Setup Approval Server

**1. Install Flask:**

```bash
pip install flask --break-system-packages
```

**2. Start the Server:**

```bash
python approval_server.py
```

This starts a web server on http://localhost:5000

**3. Make Server Public (For Production):**

**Option A: Use Ngrok (Free, Easy):**

```bash
# Download ngrok from: https://ngrok.com/download
# Then run:
ngrok http 5000
```

<!--> NB: ngrok is what i currently use, yet to deploy </!-->

This gives you a public URL like: `https://abc123.ngrok.io`

Update the `approval_server.py` line 285:

```python
server_url = "https://abc123.ngrok.io"  # Your ngrok URL
```

**Option B: Deploy to Cloud:**

- Deploy to Heroku, Railway, or Render (free tiers available)
- Or use your organization's server

**4. Update Email Function:**

The automation script will now send emails with approve/reject buttons that link to your approval server.

---

## Part 3: Complete Workflow Setup

### Fully Automated

**Run Three Processes:**

**Terminal 1: Auto-Sync (Continuous)**

```bash
python byod_automate.py  # Choose option 1 for API sync
```

Checks Google Sheets every 60 seconds

**Terminal 2: Approval Server (Always On)**

```bash
python approval_server.py
```

Handles approve/reject clicks

**Terminal 3: Make Public (If using ngrok)**

```bash
ngrok http 5000
```

## Part 4: Testing the Complete System

### Test 1: Form Submission → Auto-Sync

1. Fill out your Google Form
2. Wait 60 seconds (or run sync manually)
3. Check Excel - new row should appear
4. Check your email - confirmation sent ✓

### Test 2: Email Approval

1. Open supervisor email
2. Click "APPROVE DEVICE" button
3. Browser opens approval page
4. Click "Approve Device"
5. Check Excel - Status changed to "Approved" ✓
6. IT inspection scheduled automatically ✓

### Test 3: IT Inspection → QR Code

1. Open Excel → IT Inspection sheet
2. Fill in inspection details
3. Set Compliance Status to "Compliant"
4. Save and close
5. Wait 60 seconds (or run automation manually)
6. Check `qr_codes/` folder - new QR code created ✓
7. Check email - QR code sent ✓

---

## Part 5: Production Deployment

### Recommended Setup:

**For Small Organization (< 100 devices):**

- Use CSV sync (manual download daily)
- Run approval server on office computer
- Use ngrok for public access

**For Medium Organization (100-500 devices):**

- Use Google Sheets API auto-sync
- Deploy approval server to cloud (Heroku/Railway)
- Set up scheduled tasks

**For Large Organization (500+ devices):**

- Migrate from Excel to proper database
- Deploy full web application
- Use professional hosting

---

## Part 6: Security Considerations

**Important:**

1. **Email Credentials:** Use App Passwords, not your regular password
2. **Server Access:** Use HTTPS in production (ngrok provides this)
3. **Approval Links:** They're one-time use (could also add token expiry)
4. **Database:** Keep Excel file secure, backup regularly

**Enhanced Security (Optional):**
Add authentication to approval server:

```python
# In approval_server.py, add login requirement
from flask import session, redirect

@app.route('/approve/<reg_id>')
def approval_page(reg_id):
    if 'logged_in' not in session:
        return redirect('/login')
    # Rest of code...
```

---

## Part 7: Mobile Access

Supervisors can approve from their phones:

1. Email arrives on phone
2. Click approve button
3. Opens mobile browser
4. Review and approve
5. Done! ✓

---

## Part 8: Troubleshooting

**Auto-Sync Not Working:**

- Check credentials.json exists
- Verify Google Sheet is shared with service account
- Check internet connection
- Look for error messages in console

**Approval Buttons Not Working:**

- Make sure approval_server.py is running
- Check if ngrok URL is correct
- Verify Excel file is not open when clicking approve(this locks the systems
  out and causes an error)
- Check server console for errors

**Emails Not Sending:**

- Verify SMTP credentials
- Check spam folder
- Test with a simple email first
- Review error messages

---

## Part 9: Backup and Monitoring

**Daily Backup:**

```batch
@echo off
set BACKUP_DIR=C:\Backups\BYOD
set DATE=%date:~-4,4%%date:~-10,2%%date:~-7,2%
copy NITDA_BYOD_Database.xlsx "%BACKUP_DIR%\BYOD_%DATE%.xlsx"
```

**Monitor Logs:**

- Auto-sync shows number of new responses
- Approval server logs each approval/rejection
- Automation shows emails sent and QR codes generated

**Weekly Review:**

1. Check Dashboard sheet for statistics
2. Review any failed emails
3. Verify QR codes were generated
4. Backup database

---

## Success Checklist

- [ ] Google Form created and linked to Sheet
- [ ] Auto-sync(byod_automate.py) is running
- [ ] Approval server running(approval_server.py)
- [ ] Email approval buttons working
- [ ] IT inspection auto-scheduling working
- [ ] QR code generation working
- [ ] Backups configured
- [ ] Staff trained on the system

---

## Final Result

**Fully Automated Workflow:**

```
User fills form
    ↓ (60 seconds - automatic)
Data synced to Excel
    ↓ (automatic)
Confirmation email sent to user
    ↓ (automatic)
Approval email sent to supervisor (with buttons)
    ↓ (supervisor clicks button)
Status updated to "Approved"
    ↓ (automatic)
IT inspection scheduled
    ↓ (automatic)
Inspection notification sent
    ↓ (IT officer completes inspection)
Status updated to "Compliant"
    ↓ (automatic)
QR code generated and emailed
    ↓ (automatic)
User receives QR pass
    ↓
Security gate verification
```

**Everything runs automatically except:**

- Supervisor approval (one click)
- IT inspection (manual check)

---

## Support

If you need help:

1. Check error messages in console
2. Review this guide
3. Check TROUBLESHOOTING.md
4. Contact IT support

---

## Updates and Maintenance

**Monthly:**

- Update Python packages
- Review and update email templates
- Check for new Google Forms responses
- Verify all automations running

**Quarterly:**

- Review security settings
- Update approval workflows if needed
- Train new staff members
- Audit device compliance

---

**The End - Still buggy, just basics working, can definitely be better.**
