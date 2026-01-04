import dns.resolver
from striprtf.striprtf import rtf_to_text
import re
import sys

INPUT_FILE = "input.rtf"
OUTPUT_FILE = "output_o365.rtf"

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")

domain_cache = {}

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

def get_input_text():
    try:
        with open(INPUT_FILE, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read().strip()
            if content:
                if content.startswith("{\\rtf"):
                    return rtf_to_text(content)
                return content
    except FileNotFoundError:
        pass

    print("Paste emails below, then press CTRL+D:")
    return sys.stdin.read()

def main():
    text = get_input_text()
    emails = set(EMAIL_REGEX.findall(text))

    print(f"Found {len(emails)} emails")

    results = []

    for email in emails:
        domain = email.split("@")[1].lower()
        is_o365, provider = is_o365_domain(domain)
        if is_o365:
            results.append(f"{email}\t{domain}\t{provider}")

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(r"{\rtf1\ansi\deff0" + "\n")
        for line in results:
            f.write(line.replace("\\", r"\\") + r"\line" + "\n")
        f.write("}")

    print(f"Saved {len(results)} Outlook/O365 emails to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
