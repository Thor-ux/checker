import dns.resolver
from striprtf.striprtf import rtf_to_text
import re
import json
import time
import os
import sys

INPUT_FILE = "input.rtf"
OUTPUT_FILE = "output_o365.rtf"
CHECKPOINT_FILE = "checkpoint.json"
SAVE_EVERY = 5000

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")

# -------------------------
# Load or init checkpoint
# -------------------------
if os.path.exists(CHECKPOINT_FILE):
    with open(CHECKPOINT_FILE, "r") as f:
        checkpoint = json.load(f)
else:
    checkpoint = {
        "index": 0,
        "results": [],
        "domain_cache": {}
    }

domain_cache = checkpoint["domain_cache"]
results = checkpoint["results"]
start_index = checkpoint["index"]

# -------------------------
# O365 check
# -------------------------
def is_o365_domain(domain):
    if domain in domain_cache:
        return domain_cache[domain]

    try:
        answers = dns.resolver.resolve(domain, "MX", lifetime=5)
        for rdata in answers:
            mx = str(rdata.exchange).lower()
            if "outlook.com" in mx or "mail.protection.outlook.com" in mx:
                domain_cache[domain] = (True, mx)
                return True, mx

        domain_cache[domain] = (False, "Other")
        return False, "Other"

    except Exception:
        domain_cache[domain] = (False, "No MX")
        return False, "No MX"

# -------------------------
# Input handling
# -------------------------
def get_text():
    if os.path.exists(INPUT_FILE):
        with open(INPUT_FILE, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read().strip()
            if content.startswith("{\\rtf"):
                return rtf_to_text(content)
            return content

    print("Paste emails, then press CTRL+D")
    return sys.stdin.read()

# -------------------------
# Main processing
# -------------------------
text = get_text()
emails = list(dict.fromkeys(EMAIL_REGEX.findall(text)))
total = len(emails)

print(f"Total unique emails: {total}")
print(f"Resuming from index: {start_index}")

start_time = time.time()

for i in range(start_index, total):
    email = emails[i]
    domain = email.split("@")[1].lower()

    is_o365, provider = is_o365_domain(domain)
    if is_o365:
        results.append(f"{email}\t{domain}\t{provider}")

    # Save checkpoint
    if (i + 1) % SAVE_EVERY == 0 or i + 1 == total:
        checkpoint["index"] = i + 1
        checkpoint["results"] = results
        checkpoint["domain_cache"] = domain_cache

        with open(CHECKPOINT_FILE, "w") as f:
            json.dump(checkpoint, f)

    # Live stats
    elapsed = time.time() - start_time
    speed = int((i - start_index + 1) / elapsed) if elapsed > 0 else 0

    if (i + 1) % 1000 == 0:
        print(
            f"Processed {i+1}/{total} | "
            f"O365: {len(results)} | "
            f"{speed}/sec | "
            f"{int(elapsed)}s elapsed"
        )

# -------------------------
# Write final RTF
# -------------------------
with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write(r"{\rtf1\ansi\deff0" + "\n")
    for line in results:
        f.write(line.replace("\\", r"\\") + r"\line" + "\n")
    f.write("}")

print(f"\nDONE. Saved {len(results)} O365 emails.")
