# Polypad for OneNote — personal add-in

A small Office Add-in that puts the **Polypad** math-manipulatives canvas into a
side pane in OneNote, with a button to drop a snapshot straight onto your page.
Built for personal, non-commercial use, so it stays inside Polypad's free tier.

## What it does

- Opens Polypad (the full manipulative library) in a task pane beside your notes.
- **Insert snapshot into page** — renders the current canvas as a PNG and places it on the active OneNote page.
- **Download PNG** — reliable fallback: saves the image so you can drag it in.
- **Save / Load canvas** — copies the canvas as text you can paste into a note, then reload later to keep editing.

## Files

```
manifest.xml            the add-in definition OneNote loads
src/taskpane.html/.css/.js   the side-pane app (Polypad + buttons)
src/commands.html       required stub file referenced by the manifest
assets/icon-*.png       ribbon/button icons
```

---

## Setup — two steps

You need to (1) host these files at an HTTPS address and (2) tell OneNote to load the manifest.
The free, no-server way is **GitHub Pages**.

### Step 1 — Host the files (GitHub Pages, free)

1. Create a free account at github.com if you don't have one.
2. Make a new **public** repository named `onenote-polypad`.
3. Upload the contents of this folder (`manifest.xml`, `src/`, `assets/`) to the repo.
4. In the repo: **Settings → Pages → Build and deployment → Source: Deploy from a branch**, pick `main` / root, Save.
5. After a minute your site is live at:
   `https://YOUR-USERNAME.github.io/onenote-polypad/`
   (replace `YOUR-USERNAME` with your GitHub username).
6. Open `manifest.xml` and **find-and-replace** every
   `https://YOUR-USERNAME.github.io/onenote-polypad`
   with your real address from step 5. Re-upload the edited `manifest.xml`.

> Any HTTPS static host works (Netlify, Cloudflare Pages, your own web space).
> GitHub Pages is just the simplest free option.

### Step 2 — Load the add-in in OneNote

**OneNote on the web (easiest to test):**
1. Go to onenote.com and open a notebook.
2. **Insert → Add-ins → (Upload My Add-in)**.
3. Choose your `manifest.xml`. The **Polypad** button appears on the Home tab.

**OneNote for Windows (Microsoft 365 desktop app):**
- Desktop add-ins load from a shared-folder catalog. Easiest path:
  1. Put `manifest.xml` in a normal Windows folder, e.g. `C:\OneNoteAddins`.
  2. Share that folder (right-click → Properties → Sharing → Share) and copy its `\\PC-NAME\...` network path.
  3. In OneNote: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**, paste the network path, tick **Show in Menu**, OK, restart OneNote.
  4. **Insert → Add-ins → Shared Folder**, pick Polypad.
- Note: the **web content** (src, assets) still loads from your GitHub Pages URL — only the manifest lives in the shared folder.

Once loaded, click **Home → Polypad → Open Polypad** to launch the pane.

---

## Notes & limits

- **Free tier:** personal, non-commercial, not behind a paywall — you're covered
  (Polypad's terms allow the JS API/iframes up to ~10k requests/year in that case).
  Keep the "Polypad by Amplify" branding intact.
- **Direct insert vs. download:** OneNote is sometimes picky about images inserted
  via the API as data URIs. If **Insert snapshot** ever does nothing, use
  **Download PNG** and drag the file onto the page — same result.
- **No live editing in the note itself:** the manipulatives live in the pane; the
  page gets a picture. Use **Save canvas** to store the editable version.
- **Updating Polypad:** if Mathigon ships a newer API build, bump the version in the
  `<script src="…/polypad-en-v5.0.5.js">` line in `taskpane.html`.

## Troubleshooting

- Pane is blank / "failed to load": check the GitHub Pages URL opens in a browser,
  and that you replaced all the placeholder URLs in `manifest.xml`.
- Button doesn't appear: confirm the manifest uploaded without an error; on desktop
  confirm the Trusted Catalog path and that you restarted OneNote.
- Nothing inserts: the `ReadWriteDocument` permission is required — it's already in
  the manifest; re-upload if you edited it.
