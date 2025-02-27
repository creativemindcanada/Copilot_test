import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
import io

# Function to create the PPT presentation
def create_presentation():
    prs = Presentation()

    # Function to set an orange background for a slide
    def set_slide_background(slide, rgb_color=RGBColor(255, 165, 0)):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = rgb_color

    # ----------------------------
    # Slide 1: Why Valuezen? (Problem Statement & Opportunity)
    # ----------------------------
    slide1 = prs.slides.add_slide(prs.slide_layouts[5])
    set_slide_background(slide1)

    txBox = slide1.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(9), Inches(1.5))
    tf = txBox.text_frame
    tf.text = "Why Valuezen? (Problem Statement & Opportunity)"
    for bullet in [
        "Decision-making in logistics is slow due to lack of real-time, quantifiable value insights.",
        "Mid-market and SMBs struggle with cost justification, integration complexity, and slow adoption.",
        "Existing platforms offer data but lack personalized ROI-driven intelligence."
    ]:
        p = tf.add_paragraph()
        p.text = bullet
        p.level = 1

    # Add a logo image (if available)
    try:
        slide1.shapes.add_picture("logo.png", Inches(8), Inches(0.2), width=Inches(1.5))
    except Exception as e:
        st.warning("Logo image not found. Please ensure 'logo.png' is in your directory.")

    # ----------------------------
    # Slide 2: What is Valuezen? (Solution Overview)
    # ----------------------------
    slide2 = prs.slides.add_slide(prs.slide_layouts[5])
    set_slide_background(slide2)

    txBox = slide2.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(9), Inches(1.5))
    tf = txBox.text_frame
    tf.text = "What is Valuezen? (Solution Overview)"
    for bullet in [
        "AI-powered value delivery platform for logistics & transportation.",
        "Plug-and-play API-first approach for rapid deployment.",
        "Live value calculators to showcase impact in cost, time, and efficiency."
    ]:
        p = tf.add_paragraph()
        p.text = bullet
        p.level = 1

    try:
        slide2.shapes.add_picture("solution_image.png", Inches(0.5), Inches(3.5), width=Inches(3))
    except Exception as e:
        st.warning("Solution image not found. Please ensure 'solution_image.png' is in your directory.")

    # ----------------------------
    # Slide 3: Key Benefits (Before vs. After Valuezen)
    # ----------------------------
    slide3 = prs.slides.add_slide(prs.slide_layouts[5])
    set_slide_background(slide3)

    txBox = slide3.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1.2))
    tf = txBox.text_frame
    tf.text = "Key Benefits (Before vs. After Valuezen)"

    # Left Column: Before
    left_box = slide3.shapes.add_textbox(Inches(0.5), Inches(2), Inches(4), Inches(3))
    tf_left = left_box.text_frame
    tf_left.text = "Before:"
    for bullet in [
        "Fragmented data, unclear ROI.",
        "Slow customer onboarding & adoption.",
        "Lack of AI-driven decision intelligence."
    ]:
        p = tf_left.add_paragraph()
        p.text = bullet
        p.level = 1

    # Right Column: After
    right_box = slide3.shapes.add_textbox(Inches(5), Inches(2), Inches(4), Inches(3))
    tf_right = right_box.text_frame
    tf_right.text = "After:"
    for bullet in [
        "API-first selling → faster deployment, proven cost savings.",
        "AI/ML-driven value insights → predictive decision-making.",
        "Free-tier trials → SMB adoption & viral expansion."
    ]:
        p = tf_right.add_paragraph()
        p.text = bullet
        p.level = 1

    # ----------------------------
    # Slide 4: GTM Strategy & Expansion Plan
    # ----------------------------
    slide4 = prs.slides.add_slide(prs.slide_layouts[5])
    set_slide_background(slide4)

    txBox = slide4.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(5))
    tf = txBox.text_frame
    tf.text = "GTM Strategy & Expansion Plan"
    gtm_entries = [
        "Target Mid-Market with Rapid Deployment & API Integration",
        "  • Plug-and-play solution with API-first selling approach.",
        "  • Faster response times → reduced downtime & automation-led savings.",
        "  • Build case studies to show mid-market efficiency gains.",
        "",
        "Drive SMB Adoption with Free Tier & Simplified Onboarding",
        "  • Simplified UX → highlight ease of use & time savings in marketing.",
        "  • Free-tier tracking service for small carriers & brokers.",
        "  • Leverage referrals & integrate with popular TMS platforms.",
        "",
        "Expand into Latin America & APAC Through Mid-Market Penetration",
        "  • Localized content, multilingual support, and regional partnerships.",
        "  • Flexible pricing to match emerging market needs.",
        "",
        "Strengthen Competitive Differentiation with AI & Predictive Insights",
        "  • AI/ML-powered predictive ETA & automation as differentiators.",
        "  • Target C-level executives with data-driven supply chain intelligence.",
        "  • Position Valuezen as a decision intelligence leader."
    ]
    for entry in gtm_entries:
        p = tf.add_paragraph()
        p.text = entry
        if entry.strip().startswith("•"):
            p.level = 1
        else:
            p.level = 0

    # Save the presentation to a bytes buffer
    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer

# ----------------------------
# Streamlit App Layout
# ----------------------------
st.title("Valuezen PPT Generator")
st.write("Click the button below to generate and download your Valuezen presentation.")

ppt_data = create_presentation()

st.download_button(
    label="Download PPT",
    data=ppt_data,
    file_name="Valuezen_Presentation.pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
)

st.write("If you have any issues or see a blank screen, please check the asset paths and ensure all files are included in the repository.")
