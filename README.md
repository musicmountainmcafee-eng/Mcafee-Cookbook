# McAfee Cookbook

A recipe app: save recipes from websites, TikTok/Instagram/YouTube, photos or pasted text, keep notes on what you changed, build an ingredient list with notes and substitutions, and send a recipe to the shopping list.

Separate from the MMT dashboard: its own repo, its own Google Sheet, its own PIN.

**App:** https://musicmountainmcafee-eng.github.io/Mcafee-Cookbook/

---

## One-time setup (about 10 minutes)

### 1. Turn on GitHub Pages
Repo → **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main` / `(root)` → Save.
After a minute the app is live at the link above.

### 2. Make the Google Sheet + backend
1. Create a new blank Google Sheet. Name it something like **Cookbook Data**.
2. **Extensions → Apps Script**.
3. Delete what's in `Code.gs` and paste in everything from [`apps-script/Code.gs`](apps-script/Code.gs).
4. Click the gear (**Project Settings**) → tick **Show "appsscript.json" manifest file**. Back in the editor, open `appsscript.json` and replace it with [`apps-script/appsscript.json`](apps-script/appsscript.json).
5. Save. Pick the function **`setup`** in the toolbar and click **Run**. Approve the permissions it asks for. (Google may warn that the app isn't verified: *Advanced → Go to project*.)
6. Back in the Sheet, reload it. A **Recipe Box** menu shows up, with **Set Claude API key** (optional, see below) and **Set app PIN** (only needed to change it later).

### 3. Publish the backend
In Apps Script: **Deploy → New deployment** → type **Web app**
- Execute as: **Me**
- Who has access: **Anyone**

Deploy, then copy the **Web app URL** (ends in `/exec`).

### 4. Open the app
The backend link is built into `index.html` (`DEFAULT_URL`). The first time, type any 4 digits; they become your PIN. Wrong PINs lock the app for 15 minutes after 10 tries.
On your phone: open the link in Safari/Chrome → Share → **Add to Home Screen**.

---

## Photo & social-video imports (optional)
Website links work without anything extra: almost every recipe site publishes its recipe in a standard format the backend reads directly.

Reading **photos / screenshots**, sorting **TikTok / Instagram / YouTube captions** into ingredients and steps, and tidying **pasted text** use Claude:
1. Get an API key at https://console.anthropic.com (Settings → API Keys). It's pay-as-you-go; a recipe import costs around a cent or two.
2. In the Sheet: **Recipe Box → Set Claude API key** and paste it.

The key stays in your Apps Script settings. It's never in this repo or the app.

## Updating the backend later
Paste the new `Code.gs`, then **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**. The `/exec` link stays the same.

## What's where in the Sheet
| Tab | What's in it |
|---|---|
| Recipes | One row per recipe. `ingredients` and `steps` are readable text; the `*_data` and `original` columns on the right are what the app uses. |
| Ingredients | Your ingredient database: aisle, also-called names, notes, substitutions. Built automatically from recipes. |
| Cook Log | "Made it", adjustments, notes, and an automatic record of every edit to a recipe. |
| Shopping | The shopping list. |

## Files
- `index.html`: the whole app
- `sw.js`: lets it open offline and pick up new versions
- `manifest.json`: home-screen install + Android share menu
- `apps-script/`: the Google Sheet backend

Version tag (bottom-right of the app) uses `M.DD vN`.
