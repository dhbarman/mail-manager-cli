# mail-manager-cli

A command-line email manager using IMAP/SMTP. Currently supports Yahoo Mail with plans to extend to Gmail, Outlook, and other providers.

## Features

- Read, search, and save emails
- Send emails with attachments and HTML support
- Email templates via YAML (reusable subjects, bodies, attachments)
- Bulk delete by sender, subject, body, date — with AND/OR logic
- Clear Sent, Trash, and Spam folders
- Delete history with replay support
- Export bulk-delete criteria as Yahoo Mail filter instructions

## Setup

### 1. Install dependencies

```bash
pip install pyyaml
```

### 2. Set environment variables

**Yahoo Mail** requires an App Password (not your account password).  
Generate one at: **Yahoo → Account Security → Generate app password**

```bash
export YAHOO_EMAIL="you@yahoo.com"
export YAHOO_APP_PASSWORD="xxxx-xxxx-xxxx-xxxx"
```

Add to `~/.zshrc` or `~/.bashrc` to persist.

## Usage

### Read

```bash
python3 mail.py --inbox                        # last 20 inbox emails
python3 mail.py --inbox --limit 50             # last 50
python3 mail.py --unread                       # unread only
python3 mail.py --read UID                     # read email by UID
python3 mail.py --folders                      # list all folders
python3 mail.py --folder "Sent" --inbox        # list a specific folder
python3 mail.py --search "from:amazon"         # search by sender
python3 mail.py --search "subject:invoice"     # search by subject
```

### Save

```bash
python3 mail.py --save UID                     # save one email as .txt + .json
python3 mail.py --save-all                     # save all inbox emails
python3 mail.py --mark-read UID                # mark as read
```

### Send

```bash
# Plain text
python3 mail.py --send --to "a@b.com" --subject "Hello" --body "Hi there"

# With attachment(s)
python3 mail.py --send --to "a@b.com" --subject "Docs" --body "See attached" \
  --attach report.pdf data.csv

# HTML body
python3 mail.py --send --to "a@b.com" --subject "Update" \
  --body "<h1>Hello</h1><p>See attached.</p>" --html

# Using a template (see email_templates.yaml)
python3 mail.py --send --to "recruiter@company.com" --template job_application

# Override template subject on CLI
python3 mail.py --send --to "recruiter@company.com" --template job_application \
  --subject "Applying for Senior Engineer"

# Delete sent copy after sending
python3 mail.py --send --to "a@b.com" --template job_application --delete-sent

# List all templates
python3 mail.py --list-templates
```

### Delete

```bash
python3 mail.py --delete UID                   # delete single email by UID
python3 mail.py --clear sent                   # clear entire Sent folder
python3 mail.py --clear trash                  # clear Trash
python3 mail.py --clear spam                   # clear Spam
```

### Bulk Delete

All filter flags can be combined. Default logic is AND; use `--match-any` for OR.

```bash
# By subject keyword(s)
python3 mail.py --bulk-delete --subject-has "Newsletter"
python3 mail.py --bulk-delete --subject-has "Loan" "Insurance" "Health"

# By sender(s)
python3 mail.py --bulk-delete --from-addr "promo@example.com"
python3 mail.py --bulk-delete --from-addr "spam@a.com" "news@b.com"

# By body content
python3 mail.py --bulk-delete --body-has "click here to unsubscribe"

# By date
python3 mail.py --bulk-delete --older-than 90        # older than 90 days
python3 mail.py --bulk-delete --before 2024-01-01    # before a specific date

# Combine filters (AND — must match all)
python3 mail.py --bulk-delete --from-addr "deals@" --subject-has "offer" --older-than 30

# Combine filters (OR — match any)
python3 mail.py --bulk-delete --from-addr "promo@" --subject-has "sale" --match-any

# Preview without deleting
python3 mail.py --bulk-delete --subject-has "Health" --dry-run

# Faster body search using parallel connections
python3 mail.py --bulk-delete --body-has "unsubscribe" --parallel
```

### History & Filters

```bash
python3 mail.py --history                      # show past bulk-delete commands
python3 mail.py --replay 2                     # re-run command #2 from history
python3 mail.py --export-filters               # print Yahoo Mail filter instructions
```

## Email Templates

Define reusable templates in `email_templates.yaml`:

```yaml
job_application:
  subject: "Applying for Software Engineering roles"
  body: "Please consider the resume and cover letter attached."
  html: false
  attachments:
    - <path/to/resume>
    - <path/to/cover-letter>

invoice:
  subject: "Invoice - May 2026"
  body: "Please find the invoice attached."
  html: false
  attachments:
    - <path/to/invoice.pdf>
```

Use with `--template <name>`. CLI flags (`--subject`, `--body`, `--attach`) always override template values.

## Configuration

| Env Var | Description |
|---|---|
| `YAHOO_EMAIL` | Your Yahoo email address |
| `YAHOO_APP_PASSWORD` | Yahoo App Password (not account password) |

## File Structure

```
mail.py                  # main CLI
email_templates.yaml     # email templates
mails/                   # saved emails (auto-created, git-ignored)
delete_history.json      # bulk-delete history (git-ignored)
```

## Roadmap

- [ ] Gmail support (OAuth2)
- [ ] Outlook / Microsoft 365 support
- [ ] Attachment download from inbox
- [ ] Schedule sending
- [ ] Auto-unsubscribe detection
