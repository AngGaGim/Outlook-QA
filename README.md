# Outlook-QA: Extract Outlook Email Content and Generate QA Corpus

This project retrieves email content from an Outlook mailbox using the Microsoft Graph API and OAuth, processes it, and generates a QA corpus suitable for model training.

---

## Project Overview

- **Objective**: Extract emails from an Outlook mailbox, clean the data, and generate a QA dataset.
- **Approach**: Utilize Microsoft Graph API and OAuth protocol to access the mailbox.
- **Output**: A structured QA corpus saved in CSV format.

---

## Prerequisites

- A Microsoft account with Outlook mailbox permissions.
- An application registered in the Azure Portal to obtain an OAuth token.
- A configured Python environment with required libraries installed (`msgraph-sdk`, `requests`, etc.).

---

## Steps

### 1. Obtain Token

Microsoft no longer supports direct IMAP connections using account credentials. An OAuth token is required instead.

#### Method 1: Delegated Token Acquisition (Manual)
1. Visit [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer).
2. Sign in and generate a token with `Mail.ReadWrite` permissions.

#### Method 2: Application Token Acquisition (Programmatic)
1. Register an application in the Azure Portal.
2. Grant the application `Mail.ReadWrite` permissions.
3. Use the client credentials flow to obtain a token.

---

### 2. Email Operations

#### Method 1: Microsoft Graph API
- **Documentation**: Refer to [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/api/resources/mail-api-overview).
- **Testing**: Test the API using Graph Explorer.
- **Example**: Use the `/me/messages` endpoint to retrieve emails.

#### Method 2: Microsoft Graph SDK
- Use the Python SDK (`msgraph-sdk`).
- Sample Code: Refer to [msgraph-training-python](https://github.com/microsoftgraph/msgraph-training-python).

---

### 3. Data Processing Workflow

#### Step 1: Retrieve and Parse Emails
- Fetch raw email data using the Graph API.
- Convert the response data into JSON format.

#### Step 2: Data Filtering
- **Filtering Criteria**:
  1. Exclude single-session emails (no replies).
  2. Remove irrelevant HTML elements (tables, advertisement images, etc.).
  3. Exclude conversations with ≥ 2 messages where the last message is a forward.
  4. Process and merge forwarded emails that do not display original content.

#### Step 3: Convert to Conversation Format
- Transform email threads into a JSON conversation format.
- Use an LLM to batch-extract QA pairs.
- Save as a CSV file with columns: `Question`, `Answer`.

#### Step 4: Deduplication
- Hash QA pairs and remove duplicates.

---

## Output Format

- **File**: `qa_corpus.csv`
- **Structure**:
  ```csv
  Question,Answer
  "When is the meeting scheduled?","Tomorrow at 2 PM."
  