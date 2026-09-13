# Contact Enricher (Apify)

`enricher.py` fills the empty cells in a spreadsheet of business contacts: email, phone, website, address, revenue and employee count. It uses Apify actors and can rotate across several Apify accounts.

- **Existing cells are never overwritten.** Only empty cells are filled.
- Filled cells are highlighted **light yellow** so you can review them.
- The run can be stopped and resumed at any time.

| Step | Apify actor | Fills |
|---|---|---|
| 1 | `get-leads/linkedin-scraper` (uses your LinkedIn cookie when set) | Email, LinkedIn profile URL |
| 2a | `compass/crawler-google-places` (Google Maps, searched by company + city) | Website, phone, street, zip |
| 2b | `foxlabs/owler-intelligence` (fallback `automation-lab/owler-company-intelligence-scraper`) | Revenue, employees, and anything Google Maps missed |
| 2c | `apify/google-search-scraper` (only when the website is still unknown) | Website |
| 2d | `vdrmota/contact-info-scraper` (company home + contact pages) | Email (if it matches the person's name), business phone |
| 3 | `api-empire/linkedin-profile-phone-number-scraper` | Business phone from the LinkedIn profile (needs a LinkedIn cookie) |

---

## Quick start

```bash
pip install requests openpyxl
cp keys.example.json keys.json        # then put your Apify tokens in keys.json
python enricher.py "vivek afi.xlsx" --status     # checks your keys, no credits used
python enricher.py "vivek afi.xlsx" --dry-run    # shows what would be enriched
python enricher.py "vivek afi.xlsx" --test 3     # small real run
python enricher.py "vivek afi.xlsx" --resume     # continue with everything else
```

The rest of this guide explains each step.

---

## 1. Install

You need **Python 3.9+**.

```bash
pip install requests openpyxl
```

## 2. Get your Apify API tokens

1. Create one or more accounts at https://console.apify.com. The free tier includes $5/month of usage.
2. In each account go to **Settings → API & Integrations**.
3. Copy the **Personal API token** (it starts with `apify_api_`).
4. Open each actor from the table above on the Apify Store once and check its pricing. Some actors need a rental or a paid plan.

## 3. Put your keys in `keys.json`

`keys.json` holds your real tokens, so it is **not stored in git**. Create it from the template, in the same folder as `enricher.py`:

```bash
cp keys.example.json keys.json
```

Open `keys.json` and replace the placeholders in the `apify_keys` list. Use one entry per Apify account:

```json
"apify_keys": [
  { "token": "apify_api_YOUR_FIRST_TOKEN",  "label": "account-1", "linkedin_cookie": "li_at=AQEDAQx..." },
  { "token": "apify_api_YOUR_SECOND_TOKEN", "label": "account-2", "linkedin_cookie": null },
  { "token": "apify_api_YOUR_THIRD_TOKEN",  "label": "account-3", "linkedin_cookie": null }
]
```

| Field | What to put |
|---|---|
| `token` | The Apify API token from step 2 |
| `label` | Any unique name. It's shown in the console and logs so you can tell accounts apart |
| `linkedin_cookie` | Your LinkedIn `li_at` cookie (see step 4), or `null`. Recommended: it gives much better LinkedIn matches and enables phone lookups |

- Add or remove entries to match how many accounts you have.
- Put commas between entries, but **not** after the last one.
- Leave the rest of the file as it is. Every setting is explained in its `_comment_*` line.

To keep the file somewhere else, pass its path:

```bash
python enricher.py "vivek afi.xlsx" --keys /path/to/mykeys.json
```

Then check that the keys work. This shows each key's current Apify usage and uses no credits:

```bash
python enricher.py "vivek afi.xlsx" --status
```

## 4. (Recommended) Get your LinkedIn `li_at` cookie

The cookie is used in two places:
- **LinkedIn search (step 1).** With a cookie, the actor uses LinkedIn's own search and finds the right person far more often. Without one it falls back to search-engine results, which are shallow and often return the wrong people.
- **Phone numbers (step 3).** Phone lookups only run on keys that have a cookie.

Without any cookie you still get Google Maps and Owler data, but LinkedIn emails will be rare. To keep your LinkedIn account out of the search step, set `"linkedin_search_use_cookie": false` in `keys.json`.

> ⚠ **Use a secondary LinkedIn account, not your main one.** Scraping with your cookie is against LinkedIn's terms, and heavy use can get the account restricted. Treat the cookie like a password: anyone who has it is logged in as you.

### Chrome / Edge / Brave
1. Go to **linkedin.com** and log in.
2. Press **F12** (Mac: **Cmd + Option + I**) to open DevTools.
3. Click the **Application** tab. If you can't see it, click **»**.
4. In the left sidebar, open **Storage → Cookies → https://www.linkedin.com**.
5. Type `li_at` in the filter box.
6. Double-click the **Value** cell and copy it. It's a long string starting with `AQED...`.

### Firefox
1. Log in to linkedin.com and press **F12**.
2. Open the **Storage** tab → **Cookies** → `https://www.linkedin.com`.
3. Find `li_at` and double-click its value to copy it.

### Safari
1. Turn on the developer menu: **Safari → Settings → Advanced → "Show features for web developers"**.
2. Log in to linkedin.com and press **Cmd + Option + I**.
3. Open **Storage → Cookies**, find `li_at` and copy its value.

### Add it to `keys.json`
```json
{ "token": "apify_api_YOUR_TOKEN", "label": "account-1", "linkedin_cookie": "li_at=AQEDAQx...your value..." }
```
You can paste `li_at=AQED...` or just the `AQED...` part.

**The cookie stops working** if you log out of LinkedIn in that browser, change your password, or LinkedIn flags the session. Close the tab instead of logging out. If phone lookups start failing, copy a fresh cookie.

## 5. Run it

Run these from the folder that contains `enricher.py`.

**a) Preview.** No API calls, no credits used:
```bash
python enricher.py "vivek afi.xlsx" --dry-run
```
It lists every contact that would be processed, which fields are missing, and roughly how many actor runs it will take.

**b) Small test.** Enrich only the first 3 contacts:
```bash
python enricher.py "vivek afi.xlsx" --test 3
```
Open `enriched_vivek_afi.xlsx` and check the yellow cells before spending more credits.

**c) Everything else.** Continue from where the test stopped:
```bash
python enricher.py "vivek afi.xlsx" --resume
```

**d) Out of credits?** When every key is used up, the script saves its progress and stops with:
```
All API keys exhausted. X/244 contacts enriched. Run again with fresh keys using --resume
```
Add new tokens to `keys.json` (or wait for the monthly reset), then run `--resume` again. Keys that ran out are retried automatically after 24 hours.

**e) Check progress and key usage** at any time:
```bash
python enricher.py "vivek afi.xlsx" --status
```

Press **Ctrl+C** to stop safely. Progress is saved, and the contact that was interrupted is retried on `--resume`.

### All options

| Command | What it does |
|---|---|
| `python enricher.py input.xlsx` | Full run over all incomplete contacts |
| `--dry-run` | Show the plan. No API calls |
| `--test N` | Process only the next N incomplete contacts |
| `--rows 31-35` | Process only these Excel rows (also `A31-A35` or `31,33,40-42`). Keeps existing enriched data and doesn't change the `--resume` position |
| `--resume` | Continue from `progress.json` |
| `--restart` | Throw away previous progress and start over (overwrites the enriched file) |
| `--status` | Show progress plus live usage for each key |
| `--skip-phones` | Skip phone enrichment |
| `--keys mykeys.json` | Use a different keys file (default: `keys.json`) |
| `--delay 10` | Seconds to wait between actor calls |

If `progress.json` exists, a plain run refuses to start. Use `--resume` or `--restart`. This protects data you've already paid for.

A row counts as **incomplete** if it has a company and any of Email, Street Address, Zip Code, Business Phone, Website, Annual Revenue or Number of Employees is empty.

## Output files

All of these are written next to the input spreadsheet:

| File | Contents |
|---|---|
| `enriched_<name>.xlsx` | The enriched workbook, saved every `batch_size` contacts. `Completed?` becomes `Auto-enriched` on rows that got new data, but only where that cell was empty |
| `enrichment_log_<YYYYMMDD_HHMMSS>.txt` | Every API call, result and error. Tokens and cookies are masked |
| `progress.json` | Resume state: last row, exhausted keys, stats, company cache |

## How key rotation works

- Keys are used round-robin (`account-1 → account-2 → account-3 → account-1 …`). Set `"rotation_strategy": "sequential"` to use one key until it runs out.
- Before each actor run, the key's monthly usage is checked. Keys near `max_usage_per_key_usd` are skipped.
- A rate limit (429) or quota error marks the key **exhausted** and the next key is used. It is retried after 24 hours.
- An auth error (401/403) marks the key **dead** for the rest of the session.
- Phone lookups only use keys that have a `linkedin_cookie`. LinkedIn searches use those keys first too, and switch to cookie-less searches when they run out.

## Error handling

- Network errors and server errors are retried 3 times (after 2s, 4s, 8s).
- An actor run longer than `actor_timeout_seconds` (default 300s) is aborted, and that step is skipped for the contact. LinkedIn searches normally take about 2 minutes, so don't set this below 180.
- If an actor rejects its input, that step is turned off for the rest of the run, with a message telling you to check `keys.json`.
- A failure on one contact is logged and the run continues.

Exit codes: `0` done · `1` config/input error · `2` all keys exhausted · `130` interrupted.

## Data-quality safeguards

- A LinkedIn result is used only if the first and last name match, plus the company or the city. An email is used only when the company matches.
- Owler results must share at least one significant word with the company name.
- A Google Maps place must match the company name closely (all significant words for 1–2 word names, about two thirds for longer names) and be in the contact's state.
- The Owler street address and zip are used only when the headquarters city matches the contact's city.
- If no personal phone is found, the company's main phone number from Google Maps, Owler or the company website is used.
- A website found by web search must contain the company's first significant word in its domain (e.g. `accelergent.com` for "Accelergent Growth Solutions"). Social, news and directory sites are skipped.
- Emails from a company website are used only if they're on that website's domain and contain the person's name. Generic inboxes (`info@`, `sales@` …) are skipped unless you set `"fill_generic_company_email": true` in `keys.json`. Emails hidden by Cloudflare are decoded automatically.

## Troubleshooting

| Problem | Fix |
|---|---|
| `Keys file not found` | Run `cp keys.example.json keys.json` in the same folder as `enricher.py` |
| `has a missing or placeholder token` | Replace every `apify_api_REPLACE_...` value with a real token |
| `Invalid JSON in keys.json` | Usually a missing comma between entries or an extra one after the last. Paste the file into https://jsonlint.com |
| `INVALID TOKEN` in `--status` | The token was copied wrong or regenerated. Copy it again from Apify |
| `No keys with a linkedin_cookie` | Expected if every cookie is `null`. Phones are skipped |
| Phone lookups fail / phone step disabled | Cookie expired (get a fresh one), or the phone actor's input format changed. Edit `phone_actor_input` in `keys.json` to match the actor's Input tab on Apify |
| `shows an earlier run` | Use `--resume` to continue, or `--restart` to start over |
| `Checkpoint failed` | Close `enriched_*.xlsx` in Excel. It's retried at the next checkpoint |

## Known limitations

- `foxlabs/owler-intelligence` accepts only Owler URLs, so the script guesses the URL from the company name. When the guess is wrong, the fallback actor looks the company up by name.
- The phone actor's input format couldn't be checked because its Store page wasn't publicly reachable. Adjust `phone_actor_input` in `keys.json` if needed.
- Apify costs can't be predicted exactly. Start with `--test` and watch `--status`.
- In testing (Sep 2026), `get-leads/linkedin-scraper` returned 0 profiles for every search, even "Satya Nadella" with a valid cookie. Until that's fixed, emails mostly come from company websites. You can swap the actor ID in `keys.json`.
- The sheet can be out of date: people change jobs, so a company's contact details may no longer reach that person.
- Owler has few very small companies, so revenue and employee counts will often stay empty. Google Maps covers most local businesses for website, phone and address.
- The LinkedIn actor routes requests through a Malaysian proxy by default, and LinkedIn may block a cookie used from a different country. Set `"linkedin_proxy_country"` in `keys.json` to the country your cookie comes from (e.g. `"IN"`, `"US"`). If searches start failing, copy a fresh cookie.
