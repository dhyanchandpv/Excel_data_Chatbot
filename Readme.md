# **Excel Insight Chatbot**

**Conversational AI assistant for instant insights from Excel data using natural language.**  
Built for NeoStats AI Engineer Challenge — developed with Gemini 1.5 Flash and Streamlit.

## **Objective**

Enable non-technical users to interact with Excel spreadsheets using plain English. The chatbot reads uploaded `.xlsx` files and responds to questions with smart textual summaries, tables, or dynamic charts.

## **Key Features**

| Feature | Description |
| ----- | ----- |
| **Excel Upload** | Accepts `.xlsx` files with ≤ 500 rows and 10–20 columns |
| **Natural Queries** | Supports flexible questions like “Compare revenue by product” |
| **Schema-Agnostic** | Adapts to any tabular structure, no hardcoded column assumptions |
| **Insight Responses** | Returns clean answers as text, dataframes, or charts using Plotly |
| **Safe Code Execution** | Uses a sandboxed `exec()` to run Gemini-generated Python code |
| **Streamlined UI** | Built with Streamlit for quick deployment and intuitive user interaction |

## **Tech Stack**

* **Frontend & App**: Streamlit

* **Model**: Gemini 1.5 Flash (via Google API)

* **Data Handling**: pandas

* **Charts**: Plotly

* **Execution Safety**: Controlled Python scope

* **Hosting**: Streamlit Cloud / Hugging Face Spaces

**Core Workflow**

1. **Upload Excel** → File is read, cleaned, and column metadata extracted

2. **Ask a Question** → Natural language prompt sent to Gemini Flash

3. **Model Responds** → Generates code or direct insights

4. **Execute Code** → Securely executed with restrictions (no imports, print, I/O)

5. **Show Results** → Display answer as chart, table, or text in chat interface

