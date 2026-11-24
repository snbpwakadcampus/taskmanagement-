# SNBP Wakad – Daily Activity Command Centre

A live dashboard for the SNBP University, School of Business Management (Wakad Campus) team. Everyone can see every task, deadline and budget line. Each person logs in to update their own work. Admin logins need a one-time code (OTP) sent by email. Reminders go out by email every morning until a task is marked Completed.

**Live site (after setup):** https://snbpwakadcampus.github.io/taskmanagement-/

## What it does

| Area | Features |
|---|---|
| Dashboard (public) | Headline status, KPIs, a **deadline runway** showing each task on its due day, a team board with progress rings, overdue and due-soon lists, progress by activity, an on-time leaderboard, a live updates feed |
| My desk (login) | Your own tasks with overdue first, a progress slider, Start and ✓ Done buttons, your payments to close, and a pop-up listing pending tasks at every login |
| All tasks | Table, Kanban board (drag to change status) and calendar views, with search and filters for person, status, activity and priority |
| Budget & payments | One register for planned, approved, paid and balance amounts, payment status and due dates, linked tasks and the responsible person |
| Reports | Today, week, month or custom range; filter by person, role or status; group by day, week, month, person, role, status or activity; download as **Excel** (7 sheets) or **PDF** |
| Admin | Create logins (a welcome email goes out automatically), set access level (Admin / Manager / Member), reset passwords, bulk upload, settings and the activity log |
| Bulk upload | Excel upload for tasks, people or budget lines. Your existing *Action Plan Tracker* sheet works as is, and names in the "Responsibility" column are matched to logins |
| Reminders | Daily email at 9 AM IST for overdue tasks, tasks due soon and open payments; a team digest for admins; an email when someone is assigned a task |
| Security | OTP for admin login, salted and hashed passwords, a 15-minute lock after 6 failed tries, 6-hour sessions, forced password change on first login, and a full activity log |
| Extras | TV mode for an office screen, a "How to use" tour after login, completion confetti, auto-refresh every 90 s, mobile layout |

## Setup (about 10 minutes, one time)

### 1. Backend: Google Sheet + Apps Script (free)
1. Sign in as the Gmail account that should send the emails. Create a new **Google Sheet** named *SNBP Wakad Task Manager*.
2. **Extensions → Apps Script**. Delete the sample code, paste the whole of `backend/Code.gs`, and **Save**.
3. At the top of the code, check `CONFIG`:
   - `ADMIN_EMAIL`: your admin email. Your OTP and the temporary password go here.
   - `DASHBOARD_URL`: the GitHub Pages link above.
4. Choose the function **`setup`** in the toolbar, click **Run**, then **Allow** the permissions.
   - This creates the sheets (Users, Tasks, Budget, Settings, Log), your admin login and the daily 9 AM reminder.
   - Your **temporary admin password is emailed to you**. It's also shown in *Execution log*.
5. **Deploy → New deployment →** type **Web app** → Execute as **Me** → Who has access **Anyone** → **Deploy**. Copy the **Web app URL**. It ends in `/exec`.

### 2. Frontend: GitHub Pages
1. Open `index.html`. Near the top of the script, paste the URL:
   ```js
   const API_URL = 'https://script.google.com/macros/s/XXXX/exec';
   ```
2. Upload `index.html`, `.nojekyll`, `README.md` and the `templates` folder to **snbpwakadcampus/taskmanagement-** (*Add file → Upload files → Commit*).
3. Repo **Settings → Pages →** Source *Deploy from a branch* → `main` / `root` → Save. In about a minute the site is live at the link above.

### 3. First use
1. Open the site and choose **Log in** with your admin email and the emailed password. Enter the OTP. Set your own password.
2. **Admin → People → + Add person** for each team member, or bulk upload the *People* sheet in `templates/SNBP_Wakad_Tasks_Bulk_Upload.xlsx`. Each person gets a login email.
3. **All tasks → Bulk upload Excel** → choose your Action Plan Tracker or the template. Check the name matching, keep "Add Budget column amounts to the budget register" ticked, and click **Import**.

If `API_URL` is left empty, the page runs in **demo mode** with sample data stored only in that browser.

## Updating the backend later
After you edit `Code.gs`, go to **Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy**. The URL stays the same.

## Limits to know
- A free Gmail account can send about **100 emails a day** from Apps Script. A Google Workspace account (for example @snbpuniversity.edu.in) can send about 1,500. The remaining quota is shown under Admin → Settings.
- The Google Sheet is the database. You can open it any time as a backup, but avoid editing its columns by hand.
