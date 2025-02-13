import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import requests
from bs4 import BeautifulSoup
from transformers import pipeline
from datetime import datetime
from typing import Dict, Optional

# Load AI model
@st.cache_resource
def load_model():
    return pipeline("text-generation", model="distilgpt2")

generator = load_model()

# Function to extract website content
def scrape_competitor_website(website_url: str) -> Optional[str]:
    try:
        if not website_url.startswith(('http://', 'https://')):
            website_url = 'https://' + website_url

        response = requests.get(website_url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")
        for element in soup(['script', 'style', 'meta', 'link', 'noscript']):
            element.decompose()

        content = "\n".join([p.text.strip() for p in soup.find_all('p') if p.text.strip()])
        return content if content else None

    except Exception as e:
        st.error(f"Error fetching website content: {str(e)}")
        return None

# Function to generate AI report
def generate_ai_competitor_report(extracted_content: str) -> Optional[str]:
    try:
        prompt = f"""
        Competitor Website Analysis:
        {extracted_content}
        
        Provide insights including:
        
        CORE STRATEGIC ELEMENTS:
        Key differentiators and innovations
        
        PROVEN VALUE LEVER IMPLEMENTATION:
        Case studies and results
        
        KEY SUCCESS FACTORS:
        Factors driving success
        
        MEASURABLE OUTCOMES:
        Quantifiable impact
        """
        
        response = generator(prompt, max_new_tokens=500, num_return_sequences=1, temperature=0.7, top_p=0.9)[0]["generated_text"]
        return response.replace(prompt, "").strip()
    except Exception as e:
        st.error(f"Error generating AI report: {str(e)}")
        return None

# Streamlit UI
st.title("Competitor Website Analysis")
st.sidebar.title("Navigation")
st.subheader("Analyze a Competitor's Website")

website_url = st.text_input("Enter Competitor Website URL")

if st.button("Generate AI-Powered Report"):
    if website_url:
        with st.spinner("Scraping competitor website..."):
            extracted_content = scrape_competitor_website(website_url)
            if extracted_content:
                with st.spinner("Generating insights..."):
                    report = generate_ai_competitor_report(extracted_content)
                    if report:
                        st.success("Analysis complete! Expand the sections below.")
                        st.markdown(report)
                        st.download_button("Download Full Report", data=report, file_name=f"competitor_analysis_{datetime.now().strftime('%Y%m%d')}.txt", mime="text/plain")
                    else:
                        st.error("Failed to generate the AI report.")
            else:
                st.error("Failed to extract website content.")
    else:
        st.warning("Please enter a valid competitor website URL.")
