# Educated Email Guesser & Undelivered Email Highlighter

This Google Apps Script project automates email generation and validation within Google Sheets. It provides two key functionalities:

- Version: 0.0.2
- @tag minimum viable product

1. **Educated Email Guessing** - Generates possible email addresses based on full names, domain names, and a predefined pattern.
2. **Undelivered Email Highlighter** - Highlights email addresses in red if they appear in a list of undelivered emails.

## Features
- Adds a **custom menu** to Google Sheets with options to:
  - **Generate Email Guesses** for all rows.
  - **Highlight Undelivered Emails** by checking against a reference sheet.
- Works dynamically with columns: `FullName`, `Domain`, `Pattern`, and `EmailGuess`.
- Supports multiple email formatting patterns (e.g., `Name.Surname`, `N.Surname`, `NSurname`, etc.).
- Handles special characters in names, ensuring valid email formatting.

## Installation
1. Open **Google Sheets**.
2. Click on **Extensions → Apps Script**.
3. Copy and paste the script files into the Apps Script editor.
4. Save the project and reload the Google Sheets file.

## Usage
### **1. Generate Email Guesses**
- Click on **Custom Scripts → Generate Email Guesses**.
- The script will process all rows and populate the `EmailGuess` column based on the selected pattern.

### **2. Highlight Undelivered Emails**
- Click on **Custom Scripts → Highlight Undelivered Emails**.
- The script will check each email in the `Europe` sheet against the `Undelivered` sheet.
- Matching emails will be highlighted in **red**.

## Column Structure
| Column Name  | Description |
|-------------|-------------|
| `FullName`  | The full name of the person |
| `Domain`  | The email domain (e.g., example.com) |
| `Pattern`  | The pattern used to generate the email |
| `EmailGuess` | The generated email address |

## Custom Email Patterns Supported
| Pattern | Example Output (John Doe, example.com) |
|---------|--------------------------------|
| `Name` | `john@example.com` |
| `Surname` | `doe@example.com` |
| `Name.Surname` | `john.doe@example.com` |
| `NameSurname` | `johndoe@example.com` |
| `Name_Surname` | `john_doe@example.com` |
| `N.Surname` | `j.doe@example.com` |
| `NSurname` | `jdoe@example.com` |
| `Name.S` | `john.d@example.com` |
| `NameS` | `johnd@example.com` |
| `N.S` | `j.d@example.com` |
| `NS` | `jd@example.com` |

## License
This project is licensed under the MIT License. Feel free to modify and improve it!

## Contributions
Pull requests are welcome! If you find an issue or want to enhance the script, please open an issue or contribute via a PR.

---
🚀 Happy Scripting!

