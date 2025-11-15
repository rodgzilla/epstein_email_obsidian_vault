#!/usr/bin/env python3
"""
Email Parser - Extract email information from text files to CSV

This script parses email text files and extracts:
- Sender (From field)
- Receiver (To field)
- Date
- Email body

It handles threaded emails, multiple formats, and cleans encoding issues.
"""

import os
import re
import csv
import sys
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from dateutil import parser as date_parser


class EmailParser:
    """Parse email text files and extract structured data"""

    # Common signature patterns to remove
    SIGNATURE_PATTERNS = [
        r'\n_{10,}.*$',  # Lines starting with many underscores
        r'\n-{10,}.*$',  # Lines starting with many dashes
        r'\n\*{10,}.*$',  # Lines starting with many asterisks
        r'\nSent from my (iPhone|iPad|Android|BlackBerry).*$',
        r'\nGet Outlook for (iOS|Android).*$',
        r'\n\s*Confidential(ity)?\s*(Notice|Statement).*$',
        r'\n\s*PRIVILEGED AND CONFIDENTIAL.*$',
        r'\n\s*NOTICE:.*?(confidential|privileged).*$',
        r'\n\s*This (email|message|communication).*?(confidential|privileged|intended).*$',
    ]

    def __init__(self, base_dir: str):
        self.base_dir = Path(base_dir)
        self.stats = {
            'total_files': 0,
            'email_files': 0,
            'non_email_files': 0,
            'emails_extracted': 0,
            'errors': 0
        }

    def is_email_file(self, content: str) -> bool:
        """Check if file contains email headers"""
        # Must have at least From and one other header
        has_from = bool(re.search(r'^From:\s*.+', content, re.MULTILINE | re.IGNORECASE))
        has_subject = bool(re.search(r'^Subject:\s*', content, re.MULTILINE | re.IGNORECASE))
        has_sent = bool(re.search(r'^Sent:\s*.+', content, re.MULTILINE | re.IGNORECASE))
        has_date = bool(re.search(r'^Date:\s*.+', content, re.MULTILINE | re.IGNORECASE))

        return has_from and (has_subject or has_sent or has_date)

    def split_threaded_emails(self, content: str) -> List[str]:
        """Split content into individual emails from a thread"""
        # Look for patterns that indicate email boundaries
        # Common patterns: "From: " at start of line, "Original Message", forwarding headers

        emails = []

        # Split on "From:" at the beginning of a line (most reliable indicator)
        # But need to be careful not to split on "From:" in the body
        parts = re.split(r'\n(?=From:\s+[^\n]+\n(?:Sent:|To:|Date:|Subject:))', content)

        # Alternative: also look for "-----Original Message-----" pattern
        if len(parts) == 1:
            parts = re.split(r'\n_{5,}+\s*Original Message\s*_{5,}+', content, flags=re.IGNORECASE)

        # For each part, verify it looks like an email
        for part in parts:
            part = part.strip()
            if part and self.is_email_file(part):
                emails.append(part)

        # If no split worked, return the whole content if it's an email
        if not emails and self.is_email_file(content):
            emails.append(content)

        return emails

    def clean_name(self, name: str) -> str:
        """Clean and standardize a name or email address for use as filename"""
        if not name:
            return ''

        # Extract email from angle brackets or square brackets first
        # Pattern: "Name <email>" or "Name [email]" or "Name [mailto:email]"
        bracket_match = re.search(r'[\[<](?:mailto:)?([^\]>]+)[\]>]', name, re.IGNORECASE)
        if bracket_match:
            name = bracket_match.group(1)

        # Remove any remaining brackets and mailto: prefix
        name = re.sub(r'mailto:', '', name, flags=re.IGNORECASE)
        name = re.sub(r'[<>\[\]]+', '', name)

        # Remove leading and trailing quotes (single and double) and whitespace
        # Do this iteratively to handle nested quotes
        prev_name = None
        while prev_name != name:
            prev_name = name
            name = name.strip()
            name = name.strip('\'"')

        # Remove trailing underscores and spaces (often used for redaction)
        name = re.sub(r'[_\s]+$', '', name)

        # Remove trailing punctuation (quotes, commas, semicolons, etc.)
        name = re.sub(r'[\s;,<>\[\]\'"]+$', '', name)

        # Remove leading punctuation as well
        name = re.sub(r'^[\s;,<>\[\]\'"]+', '', name)

        # Clean up multiple spaces in the middle
        name = re.sub(r'\s+', ' ', name)

        # Final trim
        name = name.strip()

        # Normalize based on whether it's an email or name
        if '@' in name:
            # Email address: lowercase everything
            name = name.lower()
        else:
            # Name: Apply title case normalization to reduce case variations
            words = name.split()
            normalized_words = []
            for word in words:
                # Keep acronyms/initials (e.g., "E.", "Ph.D.")
                if len(word) <= 3 and word.endswith('.'):
                    normalized_words.append(word.capitalize())
                else:
                    normalized_words.append(word.capitalize())
            name = ' '.join(normalized_words)

        # Remove special characters for filesystem safety
        # Keep only: alphanumeric, spaces, dots, hyphens, underscores, @
        # Replace other special chars with empty string
        name = re.sub(r'[^\w\s.@-]', '', name)

        # Clean up any remaining multiple spaces
        name = re.sub(r'\s+', ' ', name)

        # Final trim
        name = name.strip()

        return name

    def extract_sender(self, content: str) -> str:
        """Extract sender from From: field"""
        # Look for From: header
        match = re.search(r'^From:\s*(.+?)$', content, re.MULTILINE | re.IGNORECASE)
        if not match:
            return ''

        from_line = match.group(1).strip()

        # Clean and standardize the sender name
        return self.clean_name(from_line)

    def extract_recipients(self, content: str) -> str:
        """Extract recipients from To: field"""
        # Look for To: header
        match = re.search(r'^To:\s*(.+?)(?=\n[A-Z][a-z]*:|$)', content, re.MULTILINE | re.IGNORECASE | re.DOTALL)
        if not match:
            return ''

        to_line = match.group(1).strip()

        # Handle multiple recipients separated by semicolons or commas
        cleaned_recipients = []

        # Split by semicolon or comma
        parts = re.split(r'[;,]', to_line)

        for part in parts:
            part = part.strip()
            if not part:
                continue

            # Clean each recipient using the standardization function
            cleaned = self.clean_name(part)
            if cleaned:
                cleaned_recipients.append(cleaned)

        return '; '.join(cleaned_recipients) if cleaned_recipients else self.clean_name(to_line)

    def extract_date(self, content: str) -> str:
        """Extract and normalize date from Sent: or Date: field"""
        # Try Sent: first (more common in these files)
        match = re.search(r'^Sent:\s*(.+?)$', content, re.MULTILINE | re.IGNORECASE)

        # If no Sent:, try Date:
        if not match:
            match = re.search(r'^Date:\s*(.+?)$', content, re.MULTILINE | re.IGNORECASE)

        if not match:
            return ''

        date_str = match.group(1).strip()

        # Skip if it looks like another header field (e.g., "To:", "Subject:")
        if re.match(r'^(To|From|Subject|Cc|Bcc):', date_str, re.IGNORECASE):
            return ''

        # Try to parse and standardize the date
        try:
            # Use dateutil parser which handles many formats
            # Remove timezone abbreviations that might confuse parser
            clean_date = re.sub(r'\s*\(GMT[+-]\d{2}:\d{2}\)', '', date_str)
            clean_date = re.sub(r'\s+(EDT|EST|PDT|PST|GMT|UTC|AST|CST|MST|HKT|CDT|GDT)$', '', clean_date)

            # Parse the date
            parsed_date = date_parser.parse(clean_date, fuzzy=True)

            # Return in standardized ISO format: YYYY-MM-DD HH:MM:SS
            return parsed_date.strftime('%Y-%m-%d %H:%M:%S')

        except (ValueError, TypeError, date_parser.ParserError):
            # If parsing fails, return empty string for cleaner output
            return ''

    def extract_body(self, content: str) -> str:
        """Extract email body and clean it"""
        # Find where headers end and body begins
        # Body starts after the last header field (From, To, Sent, Subject, etc.)

        # Find the last header line
        header_pattern = r'^(From|To|Sent|Date|Subject|Cc|Bcc|Importance):\s*'

        lines = content.split('\n')
        body_start_idx = 0

        for i, line in enumerate(lines):
            if re.match(header_pattern, line, re.IGNORECASE):
                body_start_idx = i + 1

        # Get body lines
        body_lines = lines[body_start_idx:]
        body = '\n'.join(body_lines).strip()

        # Clean encoding issues
        body = self.clean_encoding(body)

        # Remove signatures and disclaimers
        body = self.remove_signatures(body)

        return body.strip()

    def clean_encoding(self, text: str) -> str:
        """Fix common encoding issues"""
        # Remove BOM if present
        if text.startswith('\ufeff'):
            text = text[1:]

        # Fix common encoding issues
        replacements = {
            '\ufffd': '',  # Replacement character
            '\u00a0': ' ',  # Non-breaking space
            '\r\n': '\n',   # Normalize line endings
            '\r': '\n',
        }

        for old, new in replacements.items():
            text = text.replace(old, new)

        return text

    def remove_signatures(self, body: str) -> str:
        """Remove common email signatures and disclaimers"""
        # Apply signature patterns
        for pattern in self.SIGNATURE_PATTERNS:
            body = re.sub(pattern, '', body, flags=re.IGNORECASE | re.DOTALL)

        # Remove long confidentiality notices (often at the end)
        # Look for paragraphs with lots of legal/confidentiality keywords
        paragraphs = body.split('\n\n')
        cleaned_paragraphs = []

        confidentiality_keywords = [
            'confidential', 'privileged', 'intended recipient',
            'unauthorized', 'disclosure', 'dissemination',
            'if you are not the intended recipient'
        ]

        for para in paragraphs:
            para_lower = para.lower()
            keyword_count = sum(1 for kw in confidentiality_keywords if kw in para_lower)

            # If paragraph has 3+ confidentiality keywords, likely a disclaimer
            if keyword_count >= 3 and len(para) > 100:
                continue  # Skip this paragraph

            cleaned_paragraphs.append(para)

        return '\n\n'.join(cleaned_paragraphs).strip()

    def parse_email(self, email_content: str, filename: str) -> Optional[Dict[str, str]]:
        """Parse a single email and extract fields"""
        try:
            sender = self.extract_sender(email_content)
            receiver = self.extract_recipients(email_content)
            date = self.extract_date(email_content)
            body = self.extract_body(email_content)

            # Only return if we got at least sender or date (minimum viable email)
            if sender or date:
                return {
                    'filename': filename,
                    'sender': sender,
                    'receiver': receiver,
                    'date': date,
                    'body': body
                }

            return None

        except Exception as e:
            print(f"Error parsing email from {filename}: {e}", file=sys.stderr)
            self.stats['errors'] += 1
            return None

    def process_file(self, filepath: Path) -> List[Dict[str, str]]:
        """Process a single file and extract all emails from it"""
        self.stats['total_files'] += 1

        try:
            # Read file with encoding error handling
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                # Try with latin-1 if utf-8 fails
                with open(filepath, 'r', encoding='latin-1') as f:
                    content = f.read()

            # Check if it's an email file
            if not self.is_email_file(content):
                self.stats['non_email_files'] += 1
                return []

            self.stats['email_files'] += 1

            # Split into individual emails if threaded
            emails = self.split_threaded_emails(content)

            # Parse each email
            results = []
            for email_content in emails:
                parsed = self.parse_email(email_content, filepath.name)
                if parsed:
                    results.append(parsed)
                    self.stats['emails_extracted'] += 1

            return results

        except Exception as e:
            print(f"Error processing file {filepath}: {e}", file=sys.stderr)
            self.stats['errors'] += 1
            return []

    def parse_directory(self, output_csv: str = 'emails_extracted.csv'):
        """Parse all email files in the directory tree"""
        print(f"Starting email parsing from: {self.base_dir}")
        print(f"Output will be written to: {output_csv}\n")

        all_emails = []

        # Walk through all .txt files
        txt_files = sorted(self.base_dir.rglob('*.txt'))
        total_files = len(txt_files)

        print(f"Found {total_files} text files to process\n")

        for i, filepath in enumerate(txt_files, 1):
            if i % 100 == 0:
                print(f"Progress: {i}/{total_files} files processed "
                      f"({self.stats['emails_extracted']} emails extracted)")

            emails = self.process_file(filepath)
            all_emails.extend(emails)

        # Write to CSV
        print(f"\nWriting {len(all_emails)} emails to {output_csv}...")

        with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['filename', 'sender', 'receiver', 'date', 'body']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            writer.writeheader()
            for email in all_emails:
                writer.writerow(email)

        # Print statistics
        print("\n" + "="*60)
        print("PARSING COMPLETE")
        print("="*60)
        print(f"Total files processed:     {self.stats['total_files']}")
        print(f"Email files found:         {self.stats['email_files']}")
        print(f"Non-email files skipped:   {self.stats['non_email_files']}")
        print(f"Emails extracted:          {self.stats['emails_extracted']}")
        print(f"Errors encountered:        {self.stats['errors']}")
        print(f"Output file:               {output_csv}")
        print("="*60)


def main():
    """Main entry point"""
    # Get base directory from command line or use default
    if len(sys.argv) > 1:
        base_dir = sys.argv[1]
    else:
        # Default to TEXT directory relative to script location
        script_dir = Path(__file__).parent
        base_dir = script_dir / 'TEXT'

    # Get output file from command line or use default
    if len(sys.argv) > 2:
        output_csv = sys.argv[2]
    else:
        output_csv = 'emails_extracted.csv'

    # Verify directory exists
    if not Path(base_dir).exists():
        print(f"Error: Directory not found: {base_dir}", file=sys.stderr)
        print("\nUsage: python email_parser.py [input_directory] [output_csv]")
        print("Example: python email_parser.py TEXT emails.csv")
        sys.exit(1)

    # Create parser and run
    parser = EmailParser(base_dir)
    parser.parse_directory(output_csv)


if __name__ == '__main__':
    main()
