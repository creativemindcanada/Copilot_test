import streamlit as st
import base64
import openai
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os
import time

from dotenv import load_dotenv
load_dotenv()

# Set your OpenAI API key
openai.api_key = os.getenv('OPENAI_API_KEY')  # Change if needed

# --- Custom Formatting Settings ---
TITLE_FONT_SIZE = Pt(30)
SLIDE_FONT_SIZE = Pt(16)

# --- Functions to Generate Slide Data Using OpenAI (GPT-3) --- #
def generate_slide_titles(topic):
    prompt = f"Generate 5 slide titles for the topic '{topic}'."
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        max_tokens=200,
    )
    # Split on newline and filter out empty strings.
    return [s.strip() for s in response['choices'][0]['text'].split("\n") if s.strip() != '']

def generate_slide_content(slide_title):
    prompt = f"Generate content for the slide: '{slide_title}'."
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        max_tokens=500,  # Adjust as needed
    )
    return response['choices'][0]['text'].strip()

# --- New Function to Create a Presentation with Images and Custom Formatting --- #
def create_presentation(topic, slide_titles, slide_contents):
    prs = pptx.Presentation()
    
    # For a custom background (like an orange tone), you
    # can either modify each slide’s background via a full-slide rectangle
    # or use a PPT template (preferred if you have one).
    #
    # For example, if you have a custom template with an orange background,
    # uncomment the following line and ensure custom_template.pptx is in your folder:
    # prs = pptx.Presentation("custom_template.pptx")
    
    # -- Title Slide --
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = topic
    # Optionally, add a logo to the title slide:
    try:
        title_slide.shapes.add_picture("images/logo.png", left=Inches(8), top=Inches(0.3), width=Inches(1.5))
    except Exception as e:
        print("Logo image not found:", e)
    
    # --- Content Slides ---
    slide_layout = prs.slide_layouts[1]  # Using a title-and-content layout
    for slide_title, slide_content in zip(slide_titles, slide_contents):
        slide = prs.slides.add_slide(slide_layout)
        # Set the slide title and content.
        slide.shapes.title.text = slide_title
        try:
            # Many template layouts have a placeholder for content; adjust index if needed.
            slide.shapes.placeholders[1].text = slide_content
        except Exception as e:
            print("Error setting slide content:", e)
        
        # Customize font size for title and content.
        slide.shapes.title.text_frame.paragraphs[0].font.size = TITLE_FONT_SIZE
        # Loop through all shapes that have text frames and apply a uniform font size:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    paragraph.font.size = SLIDE_FONT_SIZE

        # Optionally, add a logo image to each slide:
        try:
            slide.shapes.add_picture("images/logo.png", left=Inches(8), top=Inches(0.3), width=Inches(1))
        except Exception as e:
            print("Logo image not found on content slide:", e)
    
    # Create a folder if it doesn't exist.
    output_dir = "generated_ppt"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    ppt_filename = f"{output_dir}/{topic}_presentation.pptx"
    prs.save(ppt_filename)
    return ppt_filename

# --- Helper Function: Make a Download Link for the PPT File ---
def get_ppt_download_link(ppt_filename):
    with open(ppt_filename, "rb") as file:
        ppt_contents = file.read()
    b64_ppt = base64.b64encode(ppt_contents).decode()
    # Use the proper MIME type for PPTX when generating download link.
    return f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64_ppt}" download="{os.path.basename(ppt_filename)}">Download the PowerPoint Presentation</a>'

# --- Main Streamlit App ---
def main():
    st.title("PowerPoint Presentation Generator with GPT-3")
    
    topic = st.text_input("Enter the topic for your presentation:")
    generate_button = st.button("Generate Presentation")
    
    if generate_button and topic:
        st.info("Generating presentation... Please wait.")
        start = time.time()
        
        slide_titles = generate_slide_titles(topic)
        st.write("Slide Titles:", slide_titles)
        
        slide_contents = [generate_slide_content(title) for title in slide_titles]
        st.write("Slide Contents:", slide_contents)
        
        ppt_filename = create_presentation(topic, slide_titles, slide_contents)
        end = time.time()
        st.write(f"Presentation generated in {round(end - start, 2)} seconds!")
        st.success("Presentation generated successfully!")
        st.markdown(get_ppt_download_link(ppt_filename), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
