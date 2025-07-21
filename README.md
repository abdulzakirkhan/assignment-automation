# RAG File Reader Service

A FastAPI-based service that processes various document types (PDF, Word, PowerPoint, Images) and extracts structured information using AI and OCR technologies. The service includes rubric detection, text extraction, and assignment draft generation capabilities.

## Features

- **Multi-format Support**: Process files in various formats:
  - Documents (PDF, DOC, DOCX)
  - Presentations (PPT, PPTX)
  - Images (JPEG, PNG, JFIF)
  - Archives (ZIP, RAR, TAR, TAR.GZ, TAR.BZ2)
 
- **Advanced Text Extraction**:
  - OCR for images using Google Cloud Vision API
  - PDF text extraction with fallback to OCR
  - PowerPoint text and image extraction
  - Word document processing with image extraction

- **AI-Powered Analysis**:
  - Rubric detection and structuring
  - Assignment information extraction
  - Draft generation capabilities
  - Smart text summarization

- **Database Integration**:
  - MySQL storage for extracted information
  - Structured data organization
  - Assignment tracking and management

## Prerequisites

### System Requirements

1. Python 3.8+
2. MySQL Server
3. System Dependencies:
   - Poppler (for PDF processing)
   - Tesseract OCR
   - UnRAR
   - LibreOffice (for document conversion)

### API Keys Required

1. OpenAI API Key
2. Google Cloud Vision API credentials

## Installation

1. Clone the repository:
   ```bash
   git clone [your-repository-url]
   cd rag-file-reader
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  
   ```

3. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Install system dependencies:

   **Ubuntu/Debian:**
   ```bash
   sudo apt-get update
   sudo apt-get install -y \
       poppler-utils \
       tesseract-ocr \
       unrar \
       libreoffice
   ```

   **macOS:**
   ```bash
   brew install \
       poppler \
       tesseract \
       unrar \
       libreoffice
   ```

   **Windows:**
   - Install [Poppler for Windows](http://blog.alivate.com.au/poppler-windows/)
   - Install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)
   - Install [WinRAR](https://www.win-rar.com/)
   - Install [LibreOffice](https://www.libreoffice.org/download/download/)

5. Set up environment variables:
   ```bash
   cp .env.example .env
   ```
   Edit `.env` file with your API keys and configuration:
   ```
   OPENAI_API_KEY=your_openai_api_key
   GOOGLE_APPLICATION_CREDENTIALS=path/to/service_account.json
   ```

## Database Setup

1. Create a MySQL database
2. Update database configuration in your `.env` file:
   ```
   DB_HOST=localhost
   DB_USER=your_username
   DB_PASSWORD=your_password
   DB_NAME=your_database_name
   ```

## Running the Service

1. Start the FastAPI server:
   ```bash
   uvicorn main:app --reload
   ```

2. Access the API documentation:
   - Swagger UI: `http://localhost:8000/docs`
   - ReDoc: `http://localhost:8000/redoc`

## API Endpoints

### POST /extract/
Process and extract information from uploaded files.

**Parameters:**
- `file`: The main document file (Required)
- `helping_material`: Additional supporting documents (Optional)
- `additional_information`: Extra context or requirements (Optional)

**Response:**
```json
{
    "data": {
        "paper_topic": "string",
        "assignment_type": "string",
        "deadline": "string",
        "word_count": "number",
        "assignment_id": "number"
        // ... additional extracted information
    },
    "message": "string",
    "status": "string"
}
```

## Error Handling

The service logs processing failures to `failed_processing_log.csv` with:
- Timestamp
- File name
- File extension
- Error message

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request


## Acknowledgments

- OpenAI for GPT models
- Google Cloud Vision for OCR capabilities
- FastAPI framework