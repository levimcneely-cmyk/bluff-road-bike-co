# Bluff Road Bike Co — backend (waiver + /admin + public stats)

One Apps Script project, bound to the private **BRBC Waiver** Google Sheet (levi@herdcoindy.com), runs:

| Piece | What it does |
|---|---|
| `waiver-backend.gs` | `/waiver` e-signatures → row in **Waivers**, signed PDF to Drive, emails |
| `admin-backend.gs` | Password-locked `/admin`: add/edit bikes, mark done, monthly profit/loss |
| GET on the same URL | Brand + style **counts only** for the About page charts (no names, emails or phones) |

The **Waivers** tab is the one list of bikes. One row = one bike. Signatures land there on their own, and bikes you add in `/admin` land there too. The **Monthly** tab holds your profit/loss per month. The old **Stats** tab is no longer needed once you're happy, so you can delete it.

## One-time setup (~10 min)

1. Open the **BRBC Waiver** Sheet → **Extensions → Apps Script**.
2. Paste `waiver-backend.gs` into `Code.gs`, replacing anything already there.
3. Click **＋ → Script**, name it `admin-backend`, and paste in `admin-backend.gs`. Save.
4. ⚙️ **Project Settings** → Time zone **America/Indiana/Indianapolis**.
5. Pick `setupAdmin` in the function dropdown → **Run** → approve the permissions. This step:
   - adds `Ref ID`, `Source` and `Charged` columns to Waivers and gives every existing bike a Ref ID,
   - creates the **Monthly** tab and copies over the Profit/Loss numbers from the old Stats tab.
6. Pick `testSignature` → **Run** once. It checks PDF and email permissions. Then delete the test row it adds.
7. Reload the Sheet → **BRBC Admin → Set password…** (10+ characters).
8. **Deploy → New deployment → Web app**. Set Execute as **Me** and Who has access **Anyone**. Copy the `/exec` URL.
9. Paste that URL into **`/config.js`** in this repo: `window.BRBC = { api: 'https://script.google.com/macros/s/…/exec' };`. Commit and push.

That one URL turns on all three: `/waiver` starts saving, `/admin` unlocks, and the About page charts read live from this Sheet. Until you set it, the About page keeps using the old stats script.

## Later

- **Changing a .gs file:** Deploy → Manage deployments → ✏️ → Version: **New version**. The URL stays the same.
- **Change password:** `/admin` → ⋯ → Change password, or in the Sheet use BRBC Admin → Set password…. Either one signs out other devices.
- **Lost phone:** BRBC Admin → Sign out all devices.
- **10 wrong passwords** trigger a 15-minute lockout. Sessions last 12 hours, or 30 days with "Keep me signed in".
- **Adding a row in the Sheet by hand works.** It gets a Ref ID the next time `/admin` loads.
- **Brand families:** Schwinn, GT, Mongoose and Roadmaster count as **GT/Mongoose/Schwinn**. Fuji, Diamondback and Nishiki count as **Fuji/Diamondback/Nishiki**. Edit `ADM_BRAND_ALIASES` in `admin-backend.gs` to change the grouping.
