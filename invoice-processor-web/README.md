# Invoice Processor Django Application

This Django-based web application is designed to process PDF invoices uploaded by users. It extracts text and images from PDFs, processes the extracted data, and sends it to a specified webhook in JSON format. The application is built with the following stack:

- Django Framework
- Bootstrap 5 for frontend styling
- Python libraries for PDF and image processing, including pdfplumber, pytesseract, and pdf2image
- Integration with OpenAI GPT for structured data extraction
- Webhooks for sending processed data

---

## Table of Contents

1. [Features](#features)
2. [Folder Structure](#folder-structure)
3. [Setup Instructions](#setup-instructions)
4. [Environment Variables](#environment-variables)
5. [Dependencies](#dependencies)
6. [Hosting Options](#hosting-options)
    - [Railway](#railway)
    - [AWS](#aws)
    - [Azure](#azure)
    - [Heroku](#heroku)
7. [Usage Instructions](#usage-instructions)
8. [Contributing](#contributing)
9. [License](#license)

---

## Features

- Upload multiple PDF files for processing.
- Drag-and-drop support for PDF uploads.
- Extract both text and images from uploaded PDFs.
- Utilize OpenAI's GPT model to format extracted data as JSON.
- Send extracted data to a webhook for further processing.
- Bootstrap 5-based responsive UI.

---

## Folder Structure

```plaintext
project/
├── manage.py
├── app/
│   ├── migrations/
│   ├── static/
│   │   ├── css/
│   │   └── js/
│   ├── templates/
│   │   └── index.html
│   ├── views.py
│   ├── models.py
│   ├── forms.py
│   └── utils/
│       ├── extract_logic.py
│       └── webhook_handler.py
├── requirements.txt
├── README.md
└── .env
```

---

## Setup Instructions

### Prerequisites

- Python 3.8+ and pip installed.
- Virtual environment (e.g., venv or Anaconda).
- Install Poppler for PDF image conversion.
  - macOS: `brew install poppler`
  - Linux: `sudo apt-get install poppler-utils`
  - Windows: [Download from Poppler for Windows](http://blog.alivate.com.au/poppler-windows/).

### Steps

1. **Clone the repository:**

    ```bash
    git clone https://github.com/yourusername/your-repo.git
    cd your-repo
    ```

2. **Create and activate a virtual environment:**

    ```bash
    python -m venv venv
    source venv/bin/activate  # Linux/MacOS
    venv\Scripts\activate     # Windows
    ```

3. **Install dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

4. **Configure .env:**

    Create a `.env` file with the following variables:

    ```env
    OPENAI_API_KEY=your-openai-api-key
    MAKE_WEBHOOK_URL=your-webhook-url
    ```

5. **Run migrations and start the server:**

    ```bash
    python manage.py makemigrations
    python manage.py migrate
    python manage.py runserver
    ```

6. **Access the application** at [http://127.0.0.1:8000](http://127.0.0.1:8000).

---

## Environment Variables

- `OPENAI_API_KEY`: API key for OpenAI GPT integration.
- `MAKE_WEBHOOK_URL`: Webhook URL for sending extracted data.

---

## Dependencies

Key dependencies include:

- **Django**: Backend framework.
- **pdfplumber**: For text extraction from PDFs.
- **pytesseract**: OCR for extracting text from images.
- **pdf2image**: Convert PDF pages to images.
- **openai**: GPT integration for structured data extraction.

For a complete list, refer to `requirements.txt`.

---

## Hosting Options

### Railway

- Add a `requirements.txt` and `Procfile`.
- Use Railway's CLI or dashboard to deploy.
- Note: Ensure Poppler is installed by adding it in the Railway environment setup.

### AWS

- Use Elastic Beanstalk or EC2.
- Configure a PostgreSQL RDS instance for the database.
- Install system dependencies for Poppler and other required utilities during deployment.

### Azure

- Use Azure App Service for deployment.
- Configure Azure Storage for temporary file handling if needed.
- Install system packages, such as Poppler, using startup scripts.

### Heroku

- Add a `Procfile`:

    ```plaintext
    web: gunicorn project.wsgi
    ```

- Configure environment variables on Heroku.
- Add a Heroku buildpack for Poppler to ensure that PDF image conversion works.

---

## Usage Instructions

1. **Upload PDF files** through the UI or use drag-and-drop.
2. **Process files** and view results directly.
3. **Extracted data** is sent to the configured webhook.

---

## Contributing

Contributions are welcome! Please create a pull request or submit issues for any bugs or feature requests.

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

