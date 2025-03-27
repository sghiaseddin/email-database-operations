# Educated Email Guesser & Undelivered Email Highlighter

This Google Apps Script project automates email generation and validation within Google Sheets. It provides two key functionalities:

1. **Educated Email Guessing** - Generates possible email addresses based on full names, domain names, and a predefined pattern.
2. **Undelivered Email Highlighter** - Highlights email addresses in red if they appear in a list of undelivered emails.
3. **Listing Same-Domain Emails** - Extracts and lists emails from protected sheets within the same domain.
4. **Get Business Description** - Request for business name using Google Custom Search JSON API, and get the title and description.
5. **Get Business Name** - Request for exact business name using Ollama API on localhost.

- Version: 1.0.0
- @tag listing same domain with insights

## Features

#### 1. Email Generation

- **Function**: `generateEmailGuesses()`
- **Description**: This function generates possible email variations based on patterns, full names, and domains provided in your spreadsheet.
- **Key Points**:
  - Uses predefined patterns (e.g., "Name.Surname", "NSurname").
  - Generates emails by combining the pattern with the domain.
  - Writes the generated emails to a specified column in your sheet.

#### 2. Highlighting Undelivered Emails

- **Function**: `highlightUndeliveredEmails()`
- **Description**: This function highlights rows where the email is not present in a provided list of valid or delivered emails.
- **Key Points**:
  - Compares each email in your spreadsheet with a list of valid emails.
  - Highlights mismatched emails (undelivered emails) for easy identification.

#### 3. Listing Same-Domain Emails from Protected Sheets

- **Function**: `listEmailsSameDomainProtected()`
- **Description**: This function extracts and lists emails from protected sheets within the same domain.
- **Key Points**:
  - Accesses data from a protected sheet (e.g., "Protected Sheet" named in your spreadsheet).
  - Filters and lists emails that belong to the same domain.

#### 4. Password Hashing for Authentication

- **Function**: `hashPassword()`
- **Description**: This function securely hashes passwords using SHA-256 algorithm.
- **Key Points**:
  - Converts plain-text passwords into hashed values for secure storage.
  - Used for authenticating users accessing the script's features.

#### 5. Custom Menu Integration

- **Function**: `onOpen()`
- **Description**: This function creates a custom menu in your Google Sheet with various options to access the script's functionalities.
- **Key Points**:
  - Adds menu items like "Generate Email Guesses," "Highlight Undelivered Emails," and more.
  - Simplifies navigation and usage of the script's features.

## Installation

#### 1. Prerequisites

- Ensure you have a Google Sheet with the necessary data (e.g., full names, domains, patterns).
- Familiarize yourself with Google Apps Script basics.
- Set up any protected sheets if required by your functions.

#### 2. Getting Started

- Open your Google Sheet and go to `Extensions > App Script`.
- Copy and paste your code into the script editor.
- Save your project with a meaningful name (e.g., "Email Generator & Analyzer").
- Test each function individually to ensure they work as expected.

#### 3. Using the Custom Menu

- After deploying your script, open your Google Sheet.
- Click on the custom menu you created (e.g., "Email Tools").
- Use the options provided to generate emails, highlight undelivered emails, and more.

### Known Limitations

- The email generation function is limited to predefined patterns. If you need additional patterns, you’ll have to modify the code.
- Highlighting undelivered emails requires a valid list of delivered emails to compare against. Without this list, the function won’t work as intended.
- Password hashing is implemented but not integrated with user authentication. You’ll need to add user authentication if required.

### Best Practices

- Always handle sensitive data (like passwords) securely.
- Regularly test your script for errors and performance issues.
- Document your code thoroughly so that others can understand and maintain it.

## Usage

### **1. Generate Email Guesses**

- Click on **Custom Scripts → Generate Email Guesses**.
- The script will process all rows and populate the `EmailGuess` column based on the selected pattern.

### **2. Highlight Undelivered Emails**

- Click on **Custom Scripts → Highlight Undelivered Emails**.
- The script will check each email in the `Europe` sheet against the `Undelivered` sheet.
- Matching emails will be highlighted in **red**.

### **3. List Same-Domain Emails from Protected Sheets**
- Click on **Custom Scripts → List Same-Domain Emails**.
- This function extracts and lists emails from a protected sheet (e.g., named "Protected Sheet") within the same domain.
- The extracted emails will be displayed in a new tab or highlighted in the spreadsheet.

### **4. Password Hashing**
- Click on **Custom Scripts → Generate Password Hash**.
- A form will appear where you can input a password.
- After submitting, the script will generate a secure hash of the password (e.g., using SHA-256).
- The hashed value will be displayed in the spreadsheet or stored in a specified location.

## Column Structure

| Column Name  | Description                            |
| ------------ | -------------------------------------- |
| `FullName`   | The full name of the person            |
| `Domain`     | The email domain (e.g., example.com)   |
| `Pattern`    | The pattern used to generate the email |
| `EmailGuess` | The generated email address            |

## Custom Email Patterns Supported

| Pattern        | Example Output (John Doe, example.com) |
| -------------- | -------------------------------------- |
| `Name`         | `john@example.com`                     |
| `Surname`      | `doe@example.com`                      |
| `Name.Surname` | `john.doe@example.com`                 |
| `NameSurname`  | `johndoe@example.com`                  |
| `Name_Surname` | `john_doe@example.com`                 |
| `N.Surname`    | `j.doe@example.com`                    |
| `NSurname`     | `jdoe@example.com`                     |
| `Name.S`       | `john.d@example.com`                   |
| `NameS`        | `johnd@example.com`                    |
| `N.S`          | `j.d@example.com`                      |
| `NS`           | `jd@example.com`                       |

## License

This project is licensed under the MIT License. Feel free to modify and improve it!

## Contributions

Pull requests are welcome! If you find an issue or want to enhance the script, please open an issue or contribute via a PR.

---

🚀 Happy Scripting!
