# OR Crisis Simulator: instructor setup (keep off GitHub)

The scenario is malignant hyperthermia; students are not told until the debrief. A live monitor intraoperative crisis scenario for ANES 1123 (Oklahoma City Community College, Anesthesia Technology program). The student plays the anesthesia technologist, answers the team and earns points for each finding and action. Students need the instructor's access code to enter.

## Files
- `index.html`: the simulator (upload to GitHub).
- `README.md`: a neutral description (upload to GitHub).
- `INSTRUCTOR-SETUP.md`: this file. Do **not** upload it to GitHub.
- `apps-script/Code.gs`: the Google Apps Script that emails report cards and logs results. Do **not** upload this to GitHub; it contains the access code. Paste it into Google Apps Script instead.

## Report card setup (once)
1. Sign in to Google as bfulton.sim.director@gmail.com and open script.google.com. Click **New project**.
2. Delete the sample code, paste in all of `Code.gs`, and click **Save**.
3. In the function menu at the top, choose **setup** and click **Run**. Approve the permissions Google asks for. You will get a test email with a link to your results sheet.
4. Click **Deploy > New deployment**. Choose type **Web app**. Set **Execute as: Me** and **Who has access: Anyone**. Click **Deploy** and copy the **Web app URL**.
5. Open `index.html`, find the line `const REPORT_URL='';` and paste the URL between the quotes. Save, then upload `index.html` to GitHub.

## Hosting with GitHub Pages
Name the repository something neutral, such as `or-crisis-sim`. The repository name appears in the link students open, so avoid any name that mentions MH.

1. Open **Settings > Pages**.
2. Choose **Deploy from a branch**, select `main` and `/ (root)`, then save.
3. Share the published link and the access code with students.
