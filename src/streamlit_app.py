import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go 
import google.generativeai as genai
import os
from typing import Any
import random

MAX_ROWS = 500
MAX_COLS = 20


GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

if 'df' not in st.session_state:
    st.session_state.df = None
if 'last_response' not in st.session_state:
    st.session_state.last_response = None
if 'last_query' not in st.session_state:
    st.session_state.last_query = None
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []
if 'suggestion' not in st.session_state:
    st.session_state.suggestion = ""
if 'example_query' not in st.session_state:
    st.session_state.example_query = None


def process_excel_file(uploaded_file) -> pd.DataFrame:
    try:
        df = pd.read_excel(uploaded_file)
        if len(df) > MAX_ROWS:
            st.error(f"Too many rows. Max allowed: {MAX_ROWS}")
            return None
        if len(df.columns) > MAX_COLS:
            st.error(f"Too many columns. Max allowed: {MAX_COLS}")
            return None
        return df
    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
        return None


def generate_prompt(df: pd.DataFrame, query: str) -> str:
    column_info = "\n".join(
        f"- {col} (type: {df[col].dtype}, sample: {df[col].dropna().head(2).tolist()})"
        for col in df.columns
    )
    prompt = f"""
You are an intelligent Excel data analyst. You are working with a pandas DataFrame named `df` and answering user questions about its content.

Your behavior:
1. If the question is factual, return the **answer in plain English**.
2. If it needs calculations/statistics/visuals, return **Python code only** using `pandas` and/or `plotly.express` with variable named `result`.

Rules:
- Do NOT import anything.
- Do NOT print or explain.
- Do NOT load files.
- Only output either a Python code block or a plain text answer.

DataFrame structure:
{column_info}

Question: {query}

Answer:
"""
    return prompt.strip()


def ask_gemini(prompt: str) -> str:
    try:
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        st.error(f"Gemini Flash error: {str(e)}")
        return ""


def execute_code(code: str, df: pd.DataFrame) -> Any:
    try:
        local_vars = {'df': df, 'pd': pd, 'px': px}
        exec(code, {"__builtins__": {}}, local_vars)
        return local_vars.get('result')
    except Exception as e:
        st.error(f"Code execution failed: {str(e)}")
        return None


def format_result(result: Any) -> str:
    if result is None:
        return "No result to display."
    if isinstance(result, pd.DataFrame):
        st.markdown("####Tabular Output")
        st.dataframe(result)
        csv = result.to_csv(index=False).encode('utf-8')
        st.download_button("Download as CSV", data=csv, file_name="result.csv", mime='text/csv')
        return "Here is the tabular data you requested."
    elif isinstance(result, go.Figure):
        st.markdown("####Chart Output")
        st.plotly_chart(result, use_container_width=True)
        return "Here is the chart you requested."
    elif isinstance(result, (str, int, float, list, dict)):
        st.markdown("####Text Output")
        st.write(result)
        return str(result)
    else:
        st.warning("Assistant returned an unknown format.")
        return "Assistant returned an unknown format."

def main():
    st.set_page_config(page_title="Excel Chatbot", page_icon=None)
    st.markdown("""
        <h1 style='text-align: center;'>Excel Chatbot Assistant</h1>
        <p style='text-align: center;'>Ask questions about your Excel data using natural language.</p>
    """, unsafe_allow_html=True)

    st.markdown("----")
    st.subheader("Upload & Preview Excel")

    if st.session_state.df is None:
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'], label_visibility="collapsed")
        if uploaded_file:
            df = process_excel_file(uploaded_file)
            if df is not None:
                st.session_state.df = df
                st.session_state.uploaded_filename = uploaded_file.name
                st.success("File uploaded successfully!")
                st.rerun()
    else:
        cols = st.columns([0.9, 0.1])
        with cols[0]:
            st.markdown(
                f"<div style='color:white;font-weight:500;background-color:#1e1e1e;padding:10px 15px;border-radius:8px;'> {st.session_state.get('uploaded_filename', 'Your Excel File')}</div>",
                unsafe_allow_html=True
            )
        with cols[1]:
            if st.button("Remove File", help="Remove uploaded file and start over", key="remove_file"):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()

        if st.button("Reset File", help="Remove uploaded file and start over"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

        with st.expander("Preview Uploaded Data"):
            st.dataframe(st.session_state.df.head())

        st.markdown("----")
        st.subheader("Ask a Question")

        with st.expander("Example Questions"):
            examples = [
                "What is the average income?",
                "Show number of customers by region.",
                "Give me a bar chart of sales per category.",
                "Compare male vs female loan approval rates."
            ]
            cols = st.columns(len(examples))
            for i, ex in enumerate(examples):
                if cols[i].button(ex):
                    st.session_state.example_query = ex
                    st.rerun()

        with st.container():
            st.markdown('<div class="chat-box" style="max-height: 300px; overflow-y: auto">', unsafe_allow_html=True)
            for msg in st.session_state.chat_history:
                if msg["sender"] == "user":
                    st.markdown(f"<div style='background-color:#1a1a1a;padding:10px;border-radius:10px;margin-bottom:5px'><b>You:</b><br>{msg['text']}</div>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<div style='background-color:#262626;padding:10px;border-radius:10px;margin-bottom:10px'><b>Assistant:</b><br>{msg['text']}</div>", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with st.form("chat_form", clear_on_submit=True):
            query = st.text_input("Type your question:", placeholder="e.g. Show average loan amount by gender", label_visibility="collapsed")
            send = st.form_submit_button("Send")

        if send and query:
            st.session_state.chat_history.append({"sender": "user", "text": query})

            with st.spinner("Thinking..."):
                prompt = generate_prompt(st.session_state.df, query)
                response = ask_gemini(prompt)
                st.session_state.last_query = query
                st.session_state.last_response = response

                if response and "result =" in response and "```python" in response:
                    try:
                        raw_code = response.split("```python")[-1].split("```")[0].strip()
                        code = "\n".join(line for line in raw_code.splitlines() if not line.strip().startswith("import"))
                        result = execute_code(code, st.session_state.df)
                        feedback = format_result(result)
                        st.session_state.chat_history.append({"sender": "assistant", "text": feedback})
                    except Exception as e:
                        st.session_state.chat_history.append({"sender": "assistant", "text": f"Failed to process code: {e}"})
                else:
                    st.session_state.chat_history.append({"sender": "assistant", "text": response if response else "No answer returned."})

        if st.session_state.example_query:
            query = st.session_state.example_query
            st.session_state.example_query = None
            st.session_state.chat_history.append({"sender": "user", "text": query})

            with st.spinner("Thinking..."):
                prompt = generate_prompt(st.session_state.df, query)
                response = ask_gemini(prompt)
                st.session_state.last_query = query
                st.session_state.last_response = response

                if response and "result =" in response and "```python" in response:
                    try:
                        raw_code = response.split("```python")[-1].split("```")[0].strip()
                        code = "\n".join(line for line in raw_code.splitlines() if not line.strip().startswith("import"))
                        result = execute_code(code, st.session_state.df)
                        feedback = format_result(result)
                        st.session_state.chat_history.append({"sender": "assistant", "text": feedback})
                    except Exception as e:
                        st.session_state.chat_history.append({"sender": "assistant", "text": f"Failed to process code: {e}"})
                else:
                    st.session_state.chat_history.append({"sender": "assistant", "text": response if response else "No answer returned."})

if __name__ == "__main__":
    main()
