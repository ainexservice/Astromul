ASTROMUL WEBSITE PACKAGE
========================

Files:
- Code.gs: Google Apps Script server-side code. Uses your existing spreadsheet.
- Index.html: Responsive mobile-first ASTROMUL website with logo, favicon, and astrologer registration form.
- ASTROMUL-logo.png: Full logo image.
- favicon.png / favicon.ico: Favicon files.
- SETUP-HINDI.txt: Setup steps.

INSTALL:
1. Open https://script.google.com/
2. Open your existing Apps Script project.
3. Replace Code.gs with this package's Code.gs.
4. Add an HTML file named exactly Index and paste the contents of Index.html.
5. Save.
6. Deploy > Manage deployments > Edit > New version > Deploy.
7. Web app should execute as: Me. Set access according to who needs to register.
8. Open the /exec web app URL and submit a test registration.
9. Confirm a tab named "Astrologer Registrations" is created in your Google Sheet.

Spreadsheet ID configured:
12AU_DqW4KwlL-JkOIIHroWqpIqxptBULUOUuKUilsvY

IMPORTANT:
- The Apps Script account needs edit permission on the target spreadsheet.
- Do not share the spreadsheet itself publicly.
- The favicon is embedded in Index.html as a data URI for the Apps Script website.
- The PNG and ICO are also included separately for other hosting setups.
- This is a website + registration backend, not a full astrology booking marketplace yet.
