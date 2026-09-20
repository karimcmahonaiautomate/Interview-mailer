# Interview request mailer

Reads from a Google Sheet a list of potential interviewees, leverages a 'template.txt' file, and then sends personalised emails per row from the sources tab.


## Files

- `send_interviews.py`: the script (Python 3)

- `config.txt`: your everyday settings (sheet URL, Gmail address, your name), so you don't retype them (Not public)


- `template.txt`: the email. First line `Subject: ...` is the subject. `[UPPERCASE PLACEHOLDERS]` are filled from columns

- Google Sheet

## Spreadsheet layout

| Tab | Columns |
|---|---|
| publications | `publication_id`, `publication_name`, `publication_description` |
| articles | `publication_id`, `article_id`, 'article_name', `article_description`, `subject_line` (optional) |
| sources | `publication_id`, `article_id`, `source_name`, `source_email`, `timeframe`, `reason_for_the_request` |

A placeholder maps to a column by name: `[PUBLICATION NAME]` becomes `publication_name`,`[REASON FOR THE REQUEST]` becomes `reason_for_the_request`. You can add columns and use them in the template the same way. Tab names are not case-sensitive.

**Subject lines:** 

Write one in the `subject_line` column of the `articles` tab and it is used for every source of that article. 

It can include placeholders, e.g. `Interview request for [PUBLICATION NAME]`. 

Subject lines never contain links: `[text](url)` shows only the text and a bare URL is dropped. 

Leave it blank (or leave the column out) and the script uses the `Subject:` line at the top of `template.txt` instead. 

**Blank publication description:** 

You can leave `publication_description` empty. The shipped `template.txt` then skips the whole
sentence "I am working with [PUBLICATION NAME], [PUBLICATION DESCRIPTION]." It does this with a conditional block.


**URLs:** 

Paste a plain URL, or write `[text to use](https://example.org)`. A link attached to a whole cell also works.

(Links on only part of a cell's text are not picked up from Google Sheets; use the `[text](url)` form for those.)

Emails go out as HTML with clickable links, plus a plain-text fallback.

## Getting your data in

The script reads Google Sheets directly using a **service account**: a robot Google identity that can see *only* the sheets you share with it, and is limited to read-only. It has no access to the rest of your Google account.


## Using the script

```bash
# 1. Dry run (the default — no email sent)

Validates everything, prints each email, saves .eml files in ./preview. Sends nothing.

(The sheet URL comes from config.txt. You can also pass --sheet "<url>" to override it.)

python send_interviews.py --gmail

# 2. Send for real 

Shows the recipient list, you must type SEND, then sends with a pause between emails

python send_interviews.py --gmail --send    

# Only one publication or article

python send_interviews.py --gmail --only-publication P1 --only-article A1
```

Safeguards: any problem (bad email, unknown ids, empty field the template needs) stops everything before anything is sent.

Each successful send is written to `sent_log.csv`, and people already in the log are skipped on later runs (`--resend` overrides).

Failures are not logged, so re-running retries only those.

## Settings file (config.txt)

`config.txt` holds `SHEET_URL`, `SMTP_USER`, `FROM_NAME`, `REPLY_TO` and optionally `CREDENTIALS`, one `NAME=value` per line.

The script reads it automatically when run from the same folder (or use `--config other.txt`). 

Order of precedence: command-line flags, then settings typed in the terminal (`export SMTP_USER=...`), then `config.txt`. Passwords are never read from the file.






