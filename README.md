To set up and use the `Python-Django-Excel2Word` , follow these steps:

### Overview

`Python-Django-Excel2Word` is a Django web application designed to extract and group common data from an Excel sheet and format it into a well-organized table in Microsoft Word.
This project can be customized and is particularly useful for tasks like preparing UAT (User Acceptance Testing) certificates or similar documents.

### Step-by-Step Installation and Usage Guide

#### 1. Clone the Repository

Start by cloning the `Python-Django-Excel2Word` repository to your local machine from GitHub.

**Command:**
```bash
git clone https://github.com/eduhapi/Python-Django-Excel2Word.git
```

This command will download the repository files to your local computer.

#### 2. Navigate to the Project Directory

Change your working directory to the newly cloned project folder.

```bash
cd Python-Django-Excel2Word
```

#### 3. Set Up a Virtual Environment

Creating a virtual environment is important to isolate your project dependencies and avoid conflicts with other Python projects.

**Create a virtual environment:**

- **Windows:**
  ```bash
  python -m venv venv
  ```

- **macOS/Linux:**
  ```bash
  python3 -m venv venv
  ```

**Activate the virtual environment:**

- **Windows:**
  ```bash
  venv\Scripts\activate
  ```

- **macOS/Linux:**
  ```bash
  source venv/bin/activate
  ```

#### 4. Install the Required Dependencies

With the virtual environment activated, install the project's dependencies listed in the `requirements.txt` file.

**Command:**
```bash
pip install -r requirements.txt
```

This will install all necessary libraries and packages needed for the application to run.

#### 5. Apply Database Migrations

Django applications require applying migrations to set up the database schema. Run the following command to apply the migrations:

**Command:**
```bash
python manage.py migrate
```

This step initializes the database for the Django project.

#### 6. Run the Development Server

After setting up the database, start the Django development server to run the application locally.

**Command:**
```bash
python manage.py runserver
```

Once the server is running, you can open a web browser and navigate to `http://127.0.0.1:8000/` to access the application.

#### 7. Upload Excel Files and Generate Word Documents

- **Upload Excel Files:** Use the web interface to upload Excel files that you want to process.
- **Generate Word Documents:** The application will extract and group common data from the uploaded Excel files and format them into a well-organized table in a Microsoft Word document.

#### 8. Customize the Application

This project is designed to be easily modified. You can adjust the data extraction, grouping logic, or formatting according to your specific needs, such as preparing UAT certificates or other structured documents.


You can effectively use the `Python-Django-Excel2Word` application for automating the conversion of Excel data into formatted Word documents.
