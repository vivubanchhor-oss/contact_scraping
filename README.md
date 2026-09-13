# Contact Enricher (Apify)

`enricher.py` fills the empty cells in a spreadsheet of business contacts: email, phone, website, address, revenue and employee count. It uses Apify actors and can rotate across several Apify accounts. **Existing cells are never overwritten.** Cells it fills are highlighted in light yellow so you can review them.

| Step | Actor | Fills |
|---|---|---|
| 1 | `get-leads/linkedin-scraper` (search_profiles mode) | Email, LinkedIn URL (used by step 3) |
| 2 | `foxlabs/owler-intelligence` → fallback `automation-lab/owler-company-intelligence-scraper` | Website, street, zip, city/state, revenue, employees. Cached per company |
| 3 | `api-empire/linkedin-profile-phone-number-scraper` | Business phone (needs a LinkedIn `li_at` cookie) |

## 1. Install

Python 3.9+.

```bash
pip install requests openpyxl
```

## 2. Get Apify tokens

1. Create one or more Apify accounts at https://console.apify.com (the free tier includes $5/month of usage).
2. In each account, go to **Settings → API & Integrations** and copy the **Personal API token**.
3. Open each actor's Store page once and check its pricing. Some actors need a rental or a paid plan.

## 3. (Optional) LinkedIn cookie for phone lookups

1. Log in to linkedin.com in your browser.
2. Open DevTools → Application → Cookies → `https://www.linkedin.com` and copy the value of `li_at`.

Scraping LinkedIn with your session cookie is against LinkedIn's terms and can get the account restricted, so use a secondary account. If no key has a cookie, phone enrichment is skipped. The script does not fail.

## 4. Configure `keys.json`

Replace the `REPLACE_...` placeholders. Set `linkedin_cookie` to `null` for accounts without one. Every option is explained in the `_comment_*` fields inside the file. The important ones:

- `max_usage_per_key_usd`: a key is skipped once its monthly Apify usage gets close to this amount.
- `batch_size`: how often (in contacts) the Excel file and `progress.json` are saved.
- `phone_actor_input`: the input template for the phone actor. **Check it against that actor's Input tab on Apify.**

Keep `keys.json` private and out of version control.

## 5. Run

```bash
python enricher.py "vivek afi.xlsx" --dry-run
```

```bash
python enricher.py "vivek afi.xlsx" --test 3
```

```bash
python enricher.py "vivek afi.xlsx" --resume
```

All options:

| Command | What it does |
|---|---|
| `python enricher.py input.xlsx` | Full run over all incomplete contacts |
| `--dry-run` | Show the plan and estimated actor runs. No API calls |
| `--test N` | Process only the next N incomplete contacts |
| `--resume` | Continue from `progress.json` (also restores exhausted keys and the company cache) |
| `--restart` | Throw away previous progress and start over |
| `--status` | Show progress plus live usage for each key |
| `--skip-phones` | Skip phone enrichment |
| `--keys mykeys.json` | Use a different keys file |
| `--delay 10` | Seconds to wait between actor calls |

A row counts as **incomplete** if it has a company and any of Email, Street Address, Zip Code, Business Phone, Website, Annual Revenue or Number of Employees is empty. Rows without a name skip the LinkedIn step but still get company data.

If `progress.json` already exists, a plain run refuses to start. Use `--resume` or `--restart`. This protects data you have already paid for.

## Output files (next to the input file)

- `enriched_<name>.xlsx`: the enriched workbook, saved at every checkpoint. `Completed?` is set to `Auto-enriched` on rows that got new data, but only where that cell was empty.
- `enrichment_log_<YYYYMMDD_HHMMSS>.txt`: every API call, result and error. Tokens and cookies are masked.
- `progress.json`: resume state.

## How key rotation works

- Keys are used round-robin. Before each actor run, the key's monthly usage is checked (`/users/me/limits`).
- A 429 response, a quota/credit error or an over-budget key marks the key **exhausted** with a timestamp, and the next key is used. Exhausted keys are retried after 24 hours.
- A 401/403 auth error marks the key **dead** for the rest of the session.
- When every key is unusable, progress is saved and the script exits with code `2`. Add fresh keys and run again with `--resume`.

## Error handling

- Network errors and 5xx responses are retried 3 times (after 2s, 4s, 8s).
- If an actor run takes longer than 120s, it is aborted and that step is skipped for the contact.
- If an actor rejects its input (HTTP 400/404), that step is turned off for the rest of the run and a message tells you to check the actor ID or input in `keys.json`.
- A failure on one contact is logged and the run moves on. Ctrl+C saves progress. The interrupted contact is retried on `--resume`.

Exit codes: `0` done, `1` config/input error, `2` all keys exhausted, `130` interrupted.

## Data-quality safeguards

- A LinkedIn result is used only if the first and last name match and either the company or the city matches. An email is used only when the company matches.
- Company results must share at least one significant word with the company name.
- The Owler street address and zip are used only when the headquarters city matches the contact's city, or the contact has no city.
- If step 3 finds no phone, the company's main phone number from Owler is used.

## Known limitations

- `foxlabs/owler-intelligence` accepts only Owler URLs, so the script guesses the URL from the company name (`owler.com/company/<name>`). When the guess is wrong, the fallback actor looks the company up by name.
- The phone actor's input format could not be checked (its Store page was not publicly reachable). If Apify rejects the default input, edit `phone_actor_input` in `keys.json`.
- The script cannot predict Apify costs exactly. Use `--test` first and check `--status`.
