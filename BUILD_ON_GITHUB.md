# Building the Mac app (and Windows .exe) on GitHub – no Python needed on your computer

You need: a GitHub account and a web browser. Nothing is installed on your Mac.

## 1. Create a private repository
1. On github.com click **+** (top right) > **New repository**.
2. Name: `portfill`. Choose **Private**. Click **Create repository**.

## 2. Upload the files
1. Unzip `PortFill_repo.zip` on your Mac.
2. In your new repository click **uploading an existing file**.
3. Drag all the files from the unzipped folder into the browser (not the folder itself): `*.py`, `ui.html`,
   `requirements.txt`, `Test_Master.xlsx`, `README.md`, `BUILD_ON_GITHUB.md` and `.gitignore` if it shows up.
4. Click **Commit changes**.

## 3. Add the build instructions (a hidden folder, so do this by hand)
1. Click **Add file > Create new file**.
2. In the name box type exactly: `.github/workflows/build.yml`
   (typing the `/` creates the folders).
3. Open `build.yml.txt` from the unzipped folder in TextEdit, select all, copy, and paste it into the big box.
4. Click **Commit changes**.

## 4. Build
1. Click the **Actions** tab. If GitHub asks, click **I understand my workflows, enable them**.
2. Click **Build PortFill** on the left, then **Run workflow** (button on the right) > **Run workflow**.
3. Wait about 5–10 minutes. Three green ticks = three builds: `mac-intel`, `mac-apple-silicon`, `windows`.
4. Click the finished run. At the bottom under **Artifacts** download **PortFill-mac-intel** (for your Intel MacBook).

## 5. Open the app on your Mac
1. Unzip the download. You get `PortFill.app`. Drag it to Applications (or just double-click it where it is).
2. The first time macOS says it can't verify the app: **right-click the app > Open > Open**. Only needed once.
   (If you only see "Move to Trash", go to System Settings > Privacy & Security, scroll down and click **Open Anyway**.)

## If something fails
Open the failed run on the Actions tab, click the red step and copy the last lines of the error to me.
Note: the build first runs the self-test; if that fails, nothing is built.
