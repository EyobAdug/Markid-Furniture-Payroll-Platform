# Put it online, build the apps, and sync

Three steps. Do them in order.

## 1. Put the code on GitHub (once)
1. On github.com click **New repository**. Name it `markid-payroll`, choose **Private**, leave every checkbox empty, **Create**.
2. On your Mac, open Terminal in the unzipped `markid-payroll` folder and run (replace YOUR-USERNAME):
```
git init
git add .
git commit -m "Markid payroll v2"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/markid-payroll.git
git push -u origin main
```
   - If `git` is missing the Mac offers to install it - accept.
   - GitHub asks you to sign in. Easiest: `brew install gh` then `gh auth login`, or use the GitHub Desktop app.
3. Open the repository's **Actions** tab. The **Tests** run should turn green.

## 2. Create the online server (the "master" copy, where the admin watches progress)
Any host that runs Docker and gives you a **persistent disk** works. The repo is ready for **Render**:
1. render.com -> **New +** -> **Blueprint** -> connect your GitHub repo -> it reads `render.yaml`.
2. When asked, type a strong **ADMIN_PASSWORD**. Click **Apply**. (A plan with a persistent disk is required; without a disk the data is erased on every restart.)
3. When it is live open the address Render shows (https://...onrender.com) and log in as `admin` with that password.
4. **Users & roles**: add your accountant, checker, approver, manager, viewer (this is done online only).
5. Copy the **SYNC_TOKEN**: Render -> your service -> **Environment** -> SYNC_TOKEN. Devices need it once.
6. Back up weekly: Settings -> **Download database backup**.

Other hosts (Fly.io, Railway, a VPS): build the `Dockerfile`, mount a volume at `/data`, set `SECRET_KEY`, `SYNC_TOKEN`, `ADMIN_PASSWORD`.
Every `git push` redeploys the server automatically on Render.

## 3. Build the Windows and Mac apps (GitHub builds them for you)
1. GitHub -> **Actions** -> **Build desktop apps** -> **Run workflow**. (Or publish a release: `git tag v2.0.0 && git push --tags`.)
2. After ~10 minutes open the finished run -> **Artifacts** (or the **Releases** page for tagged builds) and download:
   - `MarkidPayroll-Windows.zip`
   - `MarkidPayroll-Mac-AppleSilicon.zip` (M1/M2/M3/M4 Macs) or `MarkidPayroll-Mac-Intel.zip` (Apple menu -> About This Mac shows the chip)
3. Install:
   - **Mac**: unzip, drag **Markid Payroll.app** to Applications. First time: right-click -> **Open** -> Open (the app is not Apple-signed).
   - **Windows**: unzip the folder anywhere, run **MarkidPayroll.exe**. If SmartScreen warns: **More info -> Run anyway** (not code-signed).
4. First run on each computer: sign in with `admin` / `admin123` (placeholder), open **Sync & devices**, enter the server address and the SYNC_TOKEN, press **Sync now**. The server's data and users arrive; then sign in with your real server login.

## Daily use
- Work normally. With no internet everything still works; the sidebar shows **"N waiting to sync"**.
- When the internet is back the app syncs by itself every minute (or press **Sync now**).
- The admin opens **Sync & devices** on the online server to see which device sent what and when, and the Dashboard to see the payroll progress.

## How sync works (read once)
- The online server is the master. Offline entries are sent in order and applied by the server with the same rules as the app, then the app downloads the latest data.
- If the server changed meanwhile (for example the period was locked, or a price was removed) the entry is **not applied**; it appears under "Could not be applied" on that device and as REJECTED in the admin's list, so it can be entered again.
- The first sync **replaces** the computer's local data with the server's, so create real data on the server first.
- Users and passwords are managed online only. Offline logins use the user list from the last sync.
- Two people changing the same worker's payment offline: the last one to sync wins (the other is visible in the audit log).
- Install new app versions on every computer when you publish a release. A computer on a different data version refuses to sync instead of corrupting data.
- Desktop data lives in `%APPDATA%\MarkidPayroll` (Windows) or `~/Library/Application Support/MarkidPayroll` (Mac).
