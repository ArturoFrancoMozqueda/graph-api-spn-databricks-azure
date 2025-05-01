# graph-api-spn-databricks-azure

# SharePointAccess

Access and manage SharePoint content seamlessly using **Microsoft Graph API** and **MSAL**.

This Python module provides a robust client to authenticate against Microsoft Azure AD and interact programmatically with SharePoint **sites**, **files**, **folders**, and **Excel workbooks**, through the **Microsoft Graph API**.

The module integrates:
- Persistent HTTP sessions (`requests.Session`)
- Secure OAuth2 authentication (`msal`)
- Robust error handling and retries
- Utility methods for common SharePoint tasks
- Full type annotations and docstrings for clarity

---

## 📚 Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Important Microsoft Graph API References](#important-microsoft-graph-api-references)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

---

## 🚀 Features

- Authenticate securely using MSAL (**Client Credentials Flow**)
- List, download, upload, and delete files/folders in SharePoint
- Manage Excel workbooks:
  - Update worksheet ranges
  - Refresh pivot tables (bulk and individual)
  - Set cell number formats
  - Clear ranges
- Handle retries and throttling (e.g., 429 errors, timeouts)
- Download files directly into **Databricks File System (DBFS)** if needed
- Retrieve most recent files and folders easily
- Wait for files to appear (polling strategy)

---

## ⚙️ Requirements

- Python 3.8+
- Azure Active Directory Application registered with:
  - API permissions: `Files.ReadWrite.All`, `Sites.ReadWrite.All`
  - Client Secret generated
- Installed libraries:
  - `msal`
  - `requests`

---

## 📦 Installation

Clone the repository and install dependencies:

```bash
git clone https://github.com/your-org/your-repo.git
cd your-repo
pip install -r requirements.txt
```

Or you can do it manuallu by

```bash
pip install msal requests
```