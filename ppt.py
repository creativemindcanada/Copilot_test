from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Create a new Presentation
prs = Presentation()

# Function to set an orange background for a slide
def set_slide_background(slide, rgb_color=RGBColor(255, 165, 0)):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb_color

# ----------------------------
# Slide 1: Why Valuezen? (Problem Statement & Opportunity)
# ----------------------------
slide1 = prs.slides.add_slide(prs.slide_layouts[5])  # Using a blank slide layout
set_slide_background(slide1)

txBox = slide1.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(9), Inches(1.5))
tf = txBox.text_frame
tf.text = "Why Valuezen? (Problem Statement & Opportunity)"
p = tf.add_paragraph()
p.text = "• Decision-making in logistics is slow due to lack of real-time, quantifiable value insights."
p.level = 1
p = tf.add_paragraph()
p.text = "• Mid-market and SMBs struggle with cost justification, integration complexity, and slow adoption."
p.level = 1
p = tf.add_paragraph()
p.text = "• Existing platforms offer data but lack personalized ROI-driven intelligence."
p.level = 1

# Add a company logo (or any additional image) at the top-right corner
try:
    _ = slide1.shapes.add_picture("logo.png", Inches(8), Inches(0.2), width=Inches(1.5))
except Exception as e:
    print("Logo image not found, please add your logo image as 'logo.png'.")

# ----------------------------
# Slide 2: What is Valuezen? (Solution Overview)
# ----------------------------
slide2 = prs.slides.add_slide(prs.slide_layouts[5])
set_slide_background(slide2)

txBox = slide2.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(9), Inches(1.5))
tf = txBox.text_frame
tf.text = "What is Valuezen? (Solution Overview)"
p = tf.add_paragraph()
p.text = "• AI-powered value delivery platform for logistics & transportation."
p.level = 1
p = tf.add_paragraph()
p.text = "• Plug-and-play API-first approach for rapid deployment."
p.level = 1
p = tf.add_paragraph()
p.text = "• Live value calculators to showcase impact in cost, time, and efficiency."
p.level = 1

# Insert an additional solution-related image (adjust path as needed)
try:
    _ = slide2.shapes.add_picture("solution_image.png", Inches(0.5), Inches(3.5), width=Inches(3))
except Exception as e:
    print("Solution image not found, please add 'solution_image.png' in your folder.")

# ----------------------------
# Slide 3: Key Benefits (Before vs. After Valuezen)
# ----------------------------
slide3 = prs.slides.add_slide(prs.slide_layouts[5])
set_slide_background(slide3)

txBox = slide3.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1.2))
tf = txBox.text_frame
tf.text = "Key Benefits (Before vs. After Valuezen)"

# Creating two textboxes for two columns
# Left Column: Before
left_box = slide3.shapes.add_textbox(Inches(0.5), Inches(2), Inches(4), Inches(3))
tf_left = left_box.text_frame
tf_left.text = "Before:"
p = tf_left.add_paragraph()
p.text = "• Fragmented data, unclear ROI."
p.level = 1
p = tf_left.add_paragraph()
p.text = "• Slow customer onboarding & adoption."
p.level = 1
p = tf_left.add_paragraph()
p.text = "• Lack of AI-driven decision intelligence."
p.level = 1

# Right Column: After
right_box = slide3.shapes.add_textbox(Inches(5), Inches(2), Inches(4), Inches(3))
tf_right = right_box.text_frame
tf_right.text = "After:"
p = tf_right.add_paragraph()
p.text = "• API-first selling → faster deployment, proven cost savings."
p.level = 1
p = tf_right.add_paragraph()
p.text = "• AI/ML-driven value insights → predictive decision-making."
p.level = 1
p = tf_right.add_paragraph()
p.text = "• Free-tier trials → SMB adoption & viral expansion."
p.level = 1

# ----------------------------
# Slide 4: GTM Strategy & Expansion Plan
# ----------------------------
slide4 = prs.slides.add_slide(prs.slide_layouts[5])
set_slide_background(slide4)

txBox = slide4.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(5))
tf = txBox.text_frame
tf.text = "GTM Strategy & Expansion Plan"

# Define the GTM strategy details
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

# Append each entry as a new paragraph in the slide
for entry in gtm_entries:
    p = tf.add_paragraph()
    p.text = entry
    # Indent details if they start with a bullet
    if entry.strip().startswith("•"):
        p.level = 1
    else:
        p.level = 0

# ----------------------------
# Save the presentation
# ----------------------------
prs.save("Valuezen_Presentation.pptx")
print("Presentation created as 'Valuezen_Presentation.pptx'.")
