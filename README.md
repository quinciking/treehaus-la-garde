# Treehaus · La Garde Hub

The guest hub for The Treehaus Company's two November 2026 retreats at La Garde, Évran (Bretagne):

- **Retreat I:** Tue 3 – Tue 10 Nov
- **Retreat II:** Thu 12 – Thu 19 Nov

Guests open the site, enter the passcode from their invitation, and can see the week's plan, their room, meals, workshops, nearby places and the shared costs. They can update their own stay, sign up to cook, add workshops and log shared expenses.

## How it works

- `index.html` is the whole site. It's served by GitHub Pages.
- Data lives in the Google Sheet **Treehaus Hub · Live data** (one tab per list: guests, meals, workshops, expenses, settings). Edits on the site write straight to the sheet, and edits in the sheet show up on the site within about 20 seconds.
- `apps-script/Code.gs` is the small Google Apps Script that connects the two. It checks the passcode before returning or saving anything, so guest names and rooms are never in this repo.
- There are two codes: a **guest passcode** to open the hub, and a **host code** (Quinci and Tristan) that unlocks trip settings on the Hosts tab.

## One-time setup

1. Open https://script.google.com and create a **New project**. Replace its contents with `apps-script/Code.gs`.
2. Near the top, fill in `GUEST_CODE_TO_SET` and `HOST_CODE_TO_SET`. Choose `setup` in the toolbar, press **Run**, and approve the permissions. This builds the tabs in the sheet and loads the starting data.
3. **Deploy → New deployment → Web app.** Execute as: *Me*. Who has access: *Anyone*. Copy the web app URL.
4. Paste that URL into `config.js` as `apiUrl` and commit.
5. In this repo: **Settings → Pages → Deploy from a branch → main / (root)**.

You can clear the two codes in the script after step 2; they're kept in the project's Script properties. To change a code later, set it again and re-run `setup` (existing data is left alone).

## Editing

- Content that rarely changes (the week plan, places nearby, packing list) is in `index.html`.
- Everything guests can change is in the sheet. Edit cells directly if you need to; keep the `id` column as it is.

Photos of places nearby are from Wikimedia Commons, credited in the hub under each photo.
