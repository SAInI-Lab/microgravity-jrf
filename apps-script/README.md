# Application backend (Google Sheet + email)

Every application submitted on the site becomes one row in a Google Sheet you own. You get an email with the full application (reply goes straight to the applicant), and the applicant gets a receipt. If someone edits and launches again, their row is overwritten and the email says "revised".

## One-time setup

1. Create a new Google Sheet, e.g. "JRF applications 2026", in the Google account that should own the data.
2. In the Sheet: **Extensions → Apps Script**. Delete the sample code, paste in `Code.gs` from this folder, and save.
3. Check `CONFIG` at the top: `notify` is where applications are emailed (currently amruthbhat@iisc.ac.in); add anyone to `cc`.
4. In the function menu pick **setup** and press **Run**. Google asks for permission (Sheets + send email as you). Accept. On "Google hasn't verified this app", choose Advanced → Go to project. It is your own script.
5. **Deploy → New deployment → type: Web app**. Execute as: **Me**. Who has access: **Anyone**. Deploy and copy the web app URL (ends in `/exec`).
6. Put that URL in `index.html`: `var SUBMIT_URL = "https://script.google.com/macros/s/.../exec";` and push.

Test: open the URL in a browser; it should show `{"ok":true,"service":"jrf-applications",...}`.

## Notes

- If the backend is unreachable, the site falls back to asking the applicant to email their application, so nothing is lost.
- Changing `Code.gs` later: **Deploy → Manage deployments → edit → Version: New version**. The URL stays the same.
- Gmail limits: about 100 emails a day on a personal account (each application sends two).
- Spam protection: a hidden honeypot field, and duplicate sends within 20 seconds are ignored.
