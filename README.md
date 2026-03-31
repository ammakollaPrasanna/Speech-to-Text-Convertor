Speech-to-Text-Convertor
This project is a Speech-to-Text Converter integrated with a Minutes of Meeting (MoM) form. Users can directly speak into their microphone, and the spoken words are automatically transcribed into text fields in the form. The system helps in creating structured MoM reports, which can then be exported as PDF or Word documents.
A complete **Speech-to-Text based Product Lifecycle Management (PLM) and Minutes of Meeting (MoM) Automation System** developed using **AngularJS, HTML, CSS, and Flask (Python)**.

This project is an intelligent web-based system designed to simplify and automate the process of documenting meetings. It captures meeting discussions through voice input, converts speech into text in real time, organizes the content into structured meeting sections, and generates a professional document automatically.

The system integrates Speech-to-Text technology with a structured Minutes of Meeting (MoM) form, allowing users to speak directly into their microphone. The spoken content is instantly transcribed and populated into relevant fields such as agenda, discussion points, decisions, and action items. This eliminates the need for manual note-taking and significantly improves efficiency and accuracy during meetings.

Once the meeting data is captured, the system formats it into a well-structured MoM report, which can be exported as a Word or PDF document. The solution is particularly useful for organizations, teams, and project managers who require consistent and professional meeting documentation.
The system reduces manual documentation effort and improves productivity in organizational meeting workflows.

🔗 **Project Link:** *(https://github.com/ammakollaPrasanna/Speech-to-Text-Convertor)*

---

## ✨ Project Overview

The **Speech-to-Text-Convertor** project is an intelligent web-based application that allows users to:

* convert spoken input into text
* fill meeting forms using voice commands
* manage discussions and action points
* store Minutes of Meeting records
* generate downloadable Word documents
* upload and transcribe recorded audio files

The project mainly focuses on **automating the MoM creation process**.

---

## 🎯 Key Features

---

### 🎤 Real-Time Speech-to-Text Conversion

The major feature of this project is **real-time voice recognition**.

The frontend uses **Web Speech API (`webkitSpeechRecognition`)** to capture user speech and automatically convert it into text.

This feature supports:

* Variant Name
* Part Name
* Subject
* Meeting Number
* Title
* Keywords
* Meeting Day
* Venue
* Members
* Discussions
* Action Points
* Innovations
* Decisions

This enables hands-free data entry during meetings.

---

### 📋 Minutes of Meeting Automation

The system automatically structures meeting data into professional sections such as:

* Meeting information
* Discussion points
* Action items
* Presentations
* Lessons learnt
* Final decisions

This makes documentation highly organized and easy to review.

---

### ➕ Dynamic Content Management

Users can dynamically add multiple meeting sections.

Supported dynamic fields:

* Discussions
* Action points
* Responsibilities
* Work center
* Innovations
* Decisions

Each item is stored in separate lists for structured access.

---

### 📄 Automatic Word Document Generation

The Flask backend automatically generates a **professional `.docx` Word document**.

Generated document includes:

* title heading
* meeting details
* discussion list
* action points table
* decisions
* timestamp

This is implemented using **python-docx**.



---

### 🔊 Audio File Upload & Transcription

Users can upload audio recordings such as:

* `.wav`
* `.mp3`

The backend processes uploaded audio using:

* `SpeechRecognition`
* `pydub`
* Google Speech API

and converts it into text.



---

### 🧠 NLP Support

The project integrates **NLTK** for text processing.

Currently supports:

* sentence tokenization
* text formatting

Future enhancement can include:

* text summarization
* keyword extraction
* auto meeting summary



---

## 🏗️ System Workflow

```text
User Speech Input
        ↓
Speech Recognition API
        ↓
Text Conversion
        ↓
AngularJS Dynamic Form Update
        ↓
MoM Section Structuring
        ↓
Flask Backend Processing
        ↓
Word Document Generation
        ↓
Download Final Document
```

---

## 🧠 Technology Stack

| Layer                | Tools Used               |
| -------------------- | ------------------------ |
| Frontend             | HTML, CSS, AngularJS     |
| Backend              | Flask, Flask-CORS        |
| Voice Recognition    | Web Speech API           |
| Audio Processing     | SpeechRecognition, pydub |
| NLP                  | NLTK                     |
| Document Export      | python-docx              |
| Programming Language | Python, JavaScript       |

---

## 📂 Project Structure

```text
Speech-to-Text-Convertor/
│
├── app.py
├── index.html
├── requirements.txt
├── README.md
```

---

## ⚙️ Installation

Install dependencies:

```bash
pip install Flask Flask-CORS python-docx SpeechRecognition pydub nltk
```

Install **ffmpeg** for audio processing support.

---

## ▶️ Run the Project

### Backend

```bash
python app.py
```

Runs at:

```text
http://127.0.0.1:5000
```

---

### Frontend

Open:

```text
index.html
```

or run local server:

```bash
python -m http.server 8000
```

Open browser:

```text
http://localhost:8000
```


### Generate Word File

```http
POST /generate-word
```

Used for exporting meeting content as `.docx`.



---

### Audio Transcription API

```http
POST /api/transcribe_uploaded_audio
```

Used for converting uploaded speech files into text.



---

## 🚀 Main Functional Modules

---

### 📁 Administration

Handles administrative operations.

---

### 📁 Management

Used for management process tracking.

---

### 📁 Project

Contains project lifecycle details.

---

### 📁 Design

Main module where MoM generation happens.

---

### 📁 Quality

Quality assurance records.

---

### 📁 Hardware

Hardware-specific lifecycle tracking.



---

## 📈 Future Enhancements

Possible upgrades include:

* database connectivity
* MySQL / MongoDB integration
* login authentication
* cloud deployment
* PDF export
* AI summarization
* meeting analytics dashboard
* email notifications
* multilingual speech recognition

---

## 🎯 Use Cases

This project can be used in:

* corporate meetings
* project review meetings
* design reviews
* educational institutions
* management systems
* PLM environments

---

## 👨‍💻 About

The **Speech-to-Text-Convertor** project aims to simplify manual meeting documentation by using **AI-based speech recognition and document automation**.

It helps organizations reduce paperwork and improve productivity.
