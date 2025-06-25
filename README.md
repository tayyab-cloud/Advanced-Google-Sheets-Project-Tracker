# Advanced-Google-Sheets-Project-Tracker
A full-featured project management application built with Google Apps Script and Google Sheets
# ✅ Advanced Project Tracker (Google Apps Script)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Technology](https://img.shields.io/badge/Technology-Google%20Apps%20Script-blue)](https://developers.google.com/apps-script)
[![Database](https://img.shields.io/badge/Database-Google%20Sheets-green)](https://www.google.com/sheets/about/)

A comprehensive, full-featured project management application built entirely within the Google Workspace ecosystem. This tool transforms a simple Google Sheet into a dynamic, multi-user web application with a robust backend, interactive UI, and powerful automation.

---

### 🎥 Live Demo

*(This GIF showcases the application's core features in action, from adding tasks to viewing the dynamic dashboard and using the powerful search filters.)*

https://raw.githubusercontent.com/tayyab-cloud/Advanced-Google-Sheets-Project-Tracker/refs/heads/main/0625.gif



---

### ✨ Core Features

This is not just a spreadsheet; it's a complete application with a rich feature set:

#### 📋 Task Management
- **Add, Edit & Delete Tasks:** A user-friendly UI for all core task operations.
- **Automatic Task IDs:** Each task is assigned a unique, sequential ID (e.g., TSK-001).
- **Status-Based Formatting:** Rows are automatically color-coded based on their status (To Do, In Progress, Done) for instant visual feedback.
- **Automatic Sorting:** The task list is always sorted by the due date, ensuring priorities are clear.

#### 👥 Team Management
- **Add, View, Edit & Delete Members:** A complete management system for team members through a clean UI.
- **Hybrid Assignee System:** Supports multiple users with the same name by relying on unique email identifiers, preventing confusion and ensuring data integrity.
- **Robust Safety Checks:** Prevents the deletion of assignees who have active (unfinished) tasks, protecting project data.
- **Full Data Synchronization:** When an assignee's name is updated, the change is automatically reflected across all their assigned tasks and the dashboard.

#### 📊 Interactive Dashboard
- **Live Metrics:** Get a real-time count of Total Tasks, To Do, In Progress, Done, and Overdue tasks.
- **Visual Charts:**
  - **Pie Chart:** Shows the overall distribution of tasks by status.
  - **Column Chart:** Displays the workload of each assignee, correctly identifying users with the same name by their unique email.
- **Persistent UI:** Charts and dashboard elements are updated, not re-created, so they retain their position and size after every refresh.

#### ⚙️ Powerful Automation
- **Automated Email Notifications:** A daily trigger sends a consolidated reminder email to each assignee with a list of their overdue tasks.
- **Task Archiving:** A one-click feature to move all "Done" tasks to a separate "Archive" sheet, keeping the main workspace clean and focused.
- **Permanent Purge:** A safe, multi-confirmation feature to permanently delete archived tasks older than a user-specified number of days (e.g., >365 days).

#### 🔍 Advanced Search & Filter
- **Dedicated Search Sidebar:** A powerful sidebar to perform complex queries.
- **Multi-Criteria Filtering:** Filter tasks by keyword, assignee, status, or priority.
- **Deep Search:** Includes an option to extend the search into the "Archive" sheet, allowing you to find any task, past or present.

---

### 🛠️ Technology & Architecture

This application leverages the power of Google Workspace as a serverless platform:

- **Backend Logic:** **Google Apps Script (JavaScript)** serves as the powerful server-side engine that handles all data processing, business logic, automation, and email services.
- **Database:** **Google Sheets** acts as a surprisingly robust and real-time database to store all `Tasks`, `Team`, and `Archive` data.
- **Frontend UI:** **HTML, CSS (Bootstrap), and Client-Side JavaScript** are used to create the modern, responsive user interface (sidebars and dialogs) that runs inside Google Sheets.
- **Core APIs:**
  - `SpreadsheetApp`: For all database (read/write) operations.
  - `HtmlService`: To create and serve the custom UI.
  - `MailApp`: To send automated email notifications.
  - `Triggers`: To schedule the daily notification function.

---

### 📂 The Sheets Explained

The application is organized into four distinct sheets, each with a specific purpose:

1.  **`Tasks` Sheet:** This is the main workspace. It contains all active and ongoing tasks. It's designed to be clean and focused on what needs to be done now.
2.  **`Team` Sheet:** This is the "source of truth" for all personnel data. It stores the unique ID, name, email, and other details for every team member.
3.  **`Dashboard` Sheet:** The reporting hub. This sheet provides a high-level, at-a-glance overview of the entire project's health through metrics and charts.
4.  **`Archive` Sheet:** This is the historical record. All completed tasks are moved here to keep the main `Tasks` sheet efficient, without losing valuable historical data.

---

### 🚀 Getting Started: Your Own Live Demo

Get your own fully functional copy of this application in just one click!

1.  **[Click Here to Make a Copy of the Project Tracker](https://docs.google.com/spreadsheets/d/1U79Uum-SYHs2xeGwulQg9Z-zr9dsL4F7xL_Z9Xdyi4Y/copy)**
2.  Once the sheet is copied to your Google Drive, open it. The `✅ Project Tracker` menu will appear at the top after a few seconds.
3.  **Grant Permissions:** The first time you use a feature that requires authorization (like sending an email or creating a trigger), Google will ask for your permission. This is normal and safe.
4.  Explore all the features through the menu!
