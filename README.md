# PhotoNotes

Phone web app for property walk-through photos. `index.html` is the whole app (GitHub Pages). Tapping Next Photo resizes the photo, stores it on the phone, and clears the screen immediately. A queue in IndexedDB uploads each item until the Apps Script confirms that client id, so a photo is not dropped when the inspector leaves the property or loses signal.

The Apps Script source is `google-apps-script.js`. It is pasted into the Sheet; it is not deployed from this repo.

## Update Apps Script (same URL)

1. Open the walk-through Google Sheet → **Extensions → Apps Script**.
2. Replace the project contents with `google-apps-script.js`.
3. Check the `CONFIG` block at the top. **Sheet tab names and the Drive folder name may need to match the current script.** This repo did not include the previous script, so those names are best-effort:
   - Rows tab: blank `ROWS_SHEET_NAME` uses the first tab (skipping the helper tabs).
   - Follow-ups: `FollowUps` (change this if the live tab is named something else).
   - Categories: `Categories`, or unique values from the rows tab.
   - Drive: a `PhotoNotes` folder, then a folder per property and a subfolder per unit. Set `DRIVE_ROOT_FOLDER_NAME` to `''` if property folders should sit directly in My Drive. Point it at the existing folder name before deploying, or new photos will land in a second tree.
4. **Deploy → Manage deployments → Edit** (pencil) on the existing deployment → **Version: New version** → Deploy. That keeps the web app URL. A new deployment would change it.

Until that deploy, the phone still delivers photos. If the live script answers `Unknown action` for `submitPhotoRow`, the queue falls back to `uploadPhoto` then `submitRows` and keeps retrying.

## What Jeff does not need to change

The page, the Settings screen, and the saved web app URL stay as they are. End Session does not delete photos still waiting on the phone.
