import streamlit as st
from openai import OpenAI
import pandas as pd
from docx import Document
import PyPDF2
import io
from pptx import Presentation

# Show title and description.
st.title("📄 Document question answering")
st.write(
    "Upload a document below and ask a question about it – GPT will answer! "
    "To use this app, you need to provide an OpenAI API key, which you can get [here](https://platform.openai.com/account/api-keys). "
)

# Ask user for their OpenAI API key via `st.text_input`.
# Alternatively, you can store the API key in `./.streamlit/secrets.toml` and access it
# via `st.secrets`, see https://docs.streamlit.io/develop/concepts/connections/secrets-management
openai_api_key = st.text_input("OpenAI API Key", type="password")
if not openai_api_key:
    st.info("Please add your OpenAI API key to continue.", icon="🗝️")
else:

    # Create an OpenAI client.
    client = OpenAI(api_key=openai_api_key)

    # Let the user upload a file via `st.file_uploader`.
uploaded_file = st.file_uploader(
    "Upload a document (.txt, .md, .doc, .docx, .xls, .xlsx, .pdf, .pptx)",
    type=("txt", "md", "doc", "docx", "xls", "xlsx", "pdf", "pptx")
)

question = st.text_area(
    "Now ask a question about the document!",
    placeholder="Can you give me a short summary?",
    disabled=not uploaded_file,
)

if uploaded_file and question:
    if uploaded_file.name.endswith(".txt") or uploaded_file.name.endswith(".md"):
        document = uploaded_file.read().decode()

    elif uploaded_file.name.endswith(".doc") or uploaded_file.name.endswith(".docx"):
        doc = Document(uploaded_file)
        document = "\n".join([para.text for para in doc.paragraphs])

    elif uploaded_file.name.endswith(".xls") or uploaded_file.name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file)
        document = df.to_string()

    elif uploaded_file.name.endswith(".pdf"):
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        document = ""
        for page in pdf_reader.pages:
            document += page.extract_text()

    elif uploaded_file.name.endswith(".pptx"):
        presentation = Presentation(uploaded_file)
        document = ""
        for slide in presentation.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    document += shape.text + "\n"

    messages = [
        {
            "role": "user",
            "content": f"Here's a document: {document} \n\n---\n\n {question}",
        }
    ]

    stream = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=messages,
        stream=True,
    )

    st.write_stream(stream)
