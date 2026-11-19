# Setup – Sangeetha's Makeup Artist Booking System

## 1. Google Sheet + backend (5 minutes)
1. Go to sheets.google.com → create a blank sheet named "Sangeetha Bookings".
2. Extensions → Apps Script. Delete the sample code, paste all of `Code.gs`, click Save.
3. Select the function `setup` in the toolbar → Run. Approve the permissions. This creates the tabs Bookings, Payments and Tithe Records.
4. Deploy → New deployment → gear icon → Web app.
   - Execute as: Me
   - Who has access: Anyone
5. Click Deploy and copy the Web app URL (ends in /exec).
   After any later change to Code.gs: Deploy → Manage deployments → Edit → New version.

## 2. Connect the website
1. Open `index.html` in a browser (or host it, step 3).
2. Settings → paste the Web app URL → Save & connect.
   The URL is kept in that browser only; all booking data stays in Google Sheets.
3. Repeat Settings on every phone or computer you use.

## 3. Publish the website (free)
- Netlify: app.netlify.com/drop → drag the folder containing index.html → you get a public link. Or
- GitHub Pages: upload index.html to a repository → Settings → Pages → deploy from main.
Open the link on your mother's phone and use "Add to Home screen".

## Sheet structure
- Bookings: one row per booking (BookingID, date, customer, phone, event, time, location, notes, total, advance, balance, Sangeetha, Rizwana, tithe, status, confirmation).
- Payments: one row per payment (PaymentID, BookingID, date, amount).
- Tithe Records: one row per booking, always 1% of the full booking amount; edits overwrite it and deletes remove it.
